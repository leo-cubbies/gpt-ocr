import os
import json
import time
import base64
from openai import OpenAI

# .envファイルがあれば読み込む (python-dotenvが入っていれば)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass # ライブラリがなければスキップ（環境変数設定済みとみなす）

# PyMuPDF (fitz) のインポート確認
try:
    import fitz
except ImportError:
    raise ImportError("PyMuPDFが見つかりません。'pip install PyMuPDF' を実行してください。")

# --- 設定 (Configuration) ---
CONFIG = {
    "model": "gpt-5",        # 利用するモデル
    "pdf_path": "Customs Invoice Sample.pdf", # 入力PDFファイル名
    "output_path": "extracted_data_202511181744.md", # 出力ファイル名
    "batch_size": 20,             # 1回に抽出する件数
    "dpi": 300,                   # 画像変換時の解像度
}

# --- システムプロンプト定義 ---
SYSTEM_PROMPT = """
あなたは航空フォワーディングの専門家です。
あらゆるフォーマットのInvoice画像から、以下の定義に従ってデータを正規化（Normalize）して抽出してください。

# 抽出項目の定義とマッピングルール

1. **Invoice No** (請求書番号)
   - 帳票上に "Inv. No", "Bill No", "Reference" 等と記載されている場合があります。
   - HAWB番号と混同しないように注意してください。

2. **Description** (品名・摘要)
   - "Description of Goods", "Particulars", "Details" 等の列を対象とします。
   - 航空運賃の場合は "Freight Charge", "A/F" 等もここに含めます。

3. **Amount** (金額)
   - "Total", "Line Total", "Amount (USD)" 等の列。
   - 通貨記号（$, ¥）は除外し、数値のみにしてください。

# 除外ルール
- "Total Amount" や "Subtotal" などの「合計行」は抽出しないでください。明細行のみが必要です。
- ページ番号やヘッダー/フッターの情報は無視してください。

Output JSON format (list of objects):
[
  {
    "Invoice No": "文字列",
    "Item No": "文字列",
    "Product Description": "文字列",
    "Origin": "文字列",
    "Quantity": 数値またはnull,
    "Unit Value": 数値またはnull,
    "Total Value": 数値またはnull
  }
]
"""

def pdf_to_base64_images(pdf_path, dpi=300):
    """PDFをページごとのBase64画像リストに変換"""
    print(f"--- 画像変換開始: {pdf_path} (DPI={dpi}) ---")
    input_images = []
    
    try:
        with fitz.open(pdf_path) as doc:
            zoom = dpi / 72
            matrix = fitz.Matrix(zoom, zoom)
            total_pages = len(doc)
            
            for i, page in enumerate(doc):
                print(f"  -> ページ {i+1}/{total_pages} をエンコード中...")
                pix = page.get_pixmap(matrix=matrix)
                img_bytes = pix.tobytes("png")
                b64_str = base64.b64encode(img_bytes).decode('utf-8')
                
                input_images.append({
                    "type": "input_image",
                    "image_url": f"data:image/png;base64,{b64_str}"
                })
        print(f"--- 変換完了: 全 {len(input_images)} ページ ---")
        return input_images
    except Exception as e:
        print(f"\nエラー: PDFの読み込みまたは変換に失敗しました: {e}")
        return []

def json_to_markdown(json_data):
    """JSONリストをMarkdownテーブルに変換"""
    if not json_data:
        return "(データなし)"
    
    # ★ ここを修正: システムプロンプトで定義したJSONのキーと完全に一致させる
    headers = ["Invoice No", "Item No", "Product Description", "Origin", "Quantity", "Unit Value", "Total Value"]
    
    # ヘッダーとセパレーター
    lines = [
        f"| No. | {' | '.join(headers)} |",
        f"| --- |{' --- |' * len(headers)}"
    ]
    
    # データ行
    for i, item in enumerate(json_data, 1):
        row = []
        for h in headers:
            value = str(item.get(h, ''))
            # ★ Product Description の改行をスペースに置換
            if h == "Product Description":
                value = value.replace('\n', ' ').replace('\r', ' ').strip()
            row.append(value)
        lines.append(f"| {i} | {' | '.join(row)} |")
        
    return "\n".join(lines) + "\n"

def main():
    print("\n=== PDF OCR抽出ツール 開始 ===")
    
    # APIキーチェック
    if "OPENAI_API_KEY" not in os.environ:
        print("エラー: 環境変数 'OPENAI_API_KEY' が設定されていません。")
        print("PowerShellで '$env:OPENAI_API_KEY=\"sk-...\"' を実行するか、.envファイルを作成してください。")
        return

    client = OpenAI()
    
    # 1. 画像変換
    image_list = pdf_to_base64_images(CONFIG["pdf_path"], CONFIG["dpi"])
    if not image_list:
        return

    all_items = []
    batch_count = 1
    # previous_response_id と thread_id を使用しないローカルセッション管理
    session_history = [] 

    print(f"\n--- 抽出ループ開始 (バッチサイズ: {CONFIG['batch_size']}) ---")

    try:
        while True:
            print(f"\n[バッチ {batch_count}] 処理中...")
            
            # 2. リクエストデータの構築
            if batch_count == 1:
                user_msg = (
                    f"全{len(image_list)}ページの画像リストです。\n"
                    f"これらから明細の「最初の{CONFIG['batch_size']}件」をJSONで抽出してください。\n"
                    "明細がない場合は `[]` を返してください。"
                )
                # 画像リスト + テキスト指示 を一つのコンテンツリストにする
                content = image_list + [{"type": "input_text", "text": f"{SYSTEM_PROMPT}\n\n---\n\n{user_msg}"}]
                session_history.append({"role": "user", "content": content})
            else:
                user_msg = f"続きの明細を「次の{CONFIG['batch_size']}件」抽出してください。なければ `[]` を返して。"
                session_history.append({"role": "user", "content": [{"type": "input_text", "text": user_msg}]})

            # 3. APIコール
            try:
                print(f"  -> APIにリクエスト送信中...", end=" ", flush=True)
                start = time.time()
                
                response = client.responses.create(
                    model=CONFIG["model"],
                    input=session_history, # ★ previous_response_id を使わないローカルセッション
                    reasoning={"effort": "minimal"} # ★ ご指定のパラメータ
                )
                print(f"完了 ({time.time() - start:.2f}秒)")

                # 結果処理
                text = response.output_text.strip()
                
                # AIの応答をローカル履歴に追加
                session_history.append({"role": "assistant", "content": [{"type": "input_text", "text": text}]})

                # 空リスト判定
                if text == "[]":
                    print("  -> 空のリストが返されました。これ以上のデータはありません。")
                    break

                # JSONパース
                # ```json ... ``` の形式で返ってくる場合があるため、除去
                clean_json_text = text.replace("```json", "").replace("```", "").strip()

                try:
                    data = json.loads(clean_json_text)
                except json.JSONDecodeError:
                    print(f"  -> エラー: JSONパース失敗。AIの応答: {clean_json_text[:200]}...") # エラーログを修正
                    break

                if not isinstance(data, list) or not data:
                    print("  -> データ形式が不正か空です。処理を終了します。")
                    break
                
                all_items.extend(data)
                print(f"  -> {len(data)} 件抽出しました。 (累積: {len(all_items)} 件)")

                # バッチサイズ未満なら終了
                if len(data) < CONFIG["batch_size"]:
                    print("  -> 取得件数がバッチサイズ未満のため、これが最終データと判断します。")
                    break
                
                batch_count += 1
                time.sleep(1) # APIレート制限への配慮

            except Exception as e:
                print(f"\n  -> API呼び出しエラー: {e}")
                break

    finally:
        # 4. 結果出力 (Markdown)
        print("\n--- 最終保存処理 ---")
        if all_items:
            print(f"合計 {len(all_items)} 件のデータを '{CONFIG['output_path']}' に保存します...")
            output_buffer = ""
            
            # Invoice No でグルーピング
            grouped = {}
            for item in all_items:
                key = item.get("Invoice No", "UNKNOWN") # Invoice No でグルーピング
                grouped.setdefault(key, []).append(item)
            
            for key, items in grouped.items():
                output_buffer += f"\n## Invoice No: {key} ({len(items)}件)\n\n" # ★ 見出しも Invoice No に変更
                output_buffer += json_to_markdown(items)
            
            try:
                with open(CONFIG["output_path"], "w", encoding="utf-8") as f:
                    f.write(output_buffer)
                print("  -> 保存成功！")
            except Exception as e:
                print(f"  -> 保存エラー: {e}")
        else:
            print("  -> 抽出されたデータがありませんでした。保存をスキップします。")

        # 5. クリーンアップは不要 (ローカルセッション管理のため)
        print("\n(ローカルセッション方式のため、サーバー削除処理は実行しません)")

    print("\n=== 処理終了 ===")

if __name__ == "__main__":
    main()