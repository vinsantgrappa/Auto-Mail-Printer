
# メール自動処理システム

## 概要

**メール自動処理システム**は、メールの自動取得と添付ファイル処理を実現するPythonアプリケーションです。  
このツールは、業務の効率化を目的として設計されており、以下の機能を備えています：

- Gmailから未読メールを自動取得。
- メール本文を解析してPDF化。
- 添付ファイル（PDF、画像、Excel、ZIPなど）を適切に処理。
- メールアドレスから送信元の会社名を特定。

---

## 主な機能

1. **メールの自動取得**
   - GmailのIMAPプロトコルを使用して未読メールを取得。
   - メール本文や添付ファイルを自動的に解析。

2. **添付ファイル処理**
   - 添付ファイルの種類に応じた処理を実行：
     - PDFファイルをそのまま保存。
     - 画像ファイル（PNG、JPEG）をPDFに変換。
     - 不要なファイル（Excel、ZIP、TIFなど）は削除。

3. **メール本文処理**
   - メール本文をWord文書に変換し、PDFとして出力。
   - 特定のフォーマットに整形して保存。

4. **会社名の特定**
   - メールアドレスから会社名を特定し、出力ファイル名に付加。

5. **エラー管理**
   - ログファイルにエラーを記録。
   - エラー発生時に適切な処理を実施。

---

## モジュール構成

### **1. メインロジック**
- メール取得と処理の全体的なフローを管理します。
- Gmailにログインし、未読メールをチェックして自動的に処理します。

### **2. 添付ファイル処理**
- PDF、画像、ZIPファイルなどを適切に処理。
- 画像ファイルをPDFに変換し、指定ディレクトリに保存します。

### **3. メール本文解析**
- メール本文を解析し、必要なデータを抽出。
- Word文書を生成し、PDFとして保存します。

### **4. ログ管理**
- ログファイルに処理内容とエラー情報を記録。
- トラブルシューティングを支援。

---

## 使用技術

- **Python**: メインロジックの構築。
- **imaplib**: Gmailからメールを取得。
- **PyPDF2**: PDFファイルの操作。
- **Pillow**: 画像ファイルを処理。
- **docx**: Word文書の生成。
- **win32com**: Wordアプリケーションを使用したPDF変換。
- **logging**: ログ管理。

---

## 使用方法

### 必要な環境

- Python 3.8以上
- 以下のライブラリをインストール：
  ```bash
  pip install pywin32 pypdf2 pillow python-docx
  ```

### 実行方法

1. Gmailのアカウント情報を環境変数に設定します：
   - `gmail_address_mail_print`（Gmailアドレス）
   - `gmail_password_mail_print`（Gmailパスワード）

2. スクリプトを実行します：
   ```bash
   python automailprinter.py
   ```

3. 未読メールが自動的に処理され、指定のディレクトリに出力されます。

---

## 注意事項

- **セキュリティ設定**  
  GmailのIMAPアクセスを有効にし、「安全性の低いアプリのアクセス」を許可してください。
- **ファイルのサイズと形式**  
  添付ファイルが大きすぎる場合、処理が遅くなる可能性があります。

---

## トラブルシューティング

### よくある問題と解決策

1. **ログインエラーが発生**  
   - GmailのIMAPアクセスが有効になっていることを確認してください。
   - 正しい環境変数が設定されているか確認してください。

2. **WordのPDF変換エラー**  
   - WindowsにMicrosoft Wordがインストールされていることを確認してください。
   - 必要に応じて、`win32com`を再インストールしてください：
     ```bash
     pip install pywin32 --force-reinstall
     ```

---

## ライセンス

本プロジェクトはMITライセンスの下で公開されています。