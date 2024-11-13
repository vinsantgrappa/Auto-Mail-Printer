```markdown
# AutoMailPrinter

AutoMailPrinterは、Gmailアカウントから未読メールを自動で取得し、添付ファイルや本文を処理し、PDFとして保存するPythonスクリプトです。指定されたフォルダに出力し、ExcelやZIPファイルなど特定のファイルを削除する機能も備えています。

## 機能
- **未読メールの取得**: Gmailから未読のメールを取得します。
- **添付ファイルの処理**: メールに添付されたPDFや画像ファイルを取得し、所定のフォルダに保存します。
- **ファイルの変換**: 画像ファイルをPDFに変換し、WordファイルもPDF形式で保存可能。
- **メール本文の保存**: メール本文をテキストやWordファイルに書き込み、さらにPDFに変換します。
- **ログ記録**: ログファイルに詳細なログ情報を保存します。

## 依存関係
このプロジェクトを実行するには、以下のPythonパッケージが必要です：
- `PyPDF2`
- `PIL` (Pillow)
- `python-docx`
- `win32com.client`
- `imaplib`
- `email`

以下のコマンドで依存関係をインストールできます：
```bash
pip install PyPDF2 pillow python-docx pypiwin32
```

## 環境変数の設定
環境変数にGmailの認証情報を設定する必要があります：
- `gmail_address_mail_print`: Gmailアカウントのメールアドレス
- `gmail_password_mail_print`: Gmailアカウントのパスワード

## 使用方法
1. 必要なPythonパッケージをインストールします。
2. `AutoMailPrinter` クラスを実行して、未読メールを自動で取得・処理します。
3. `run()` メソッドは15秒毎に新しい未読メールをチェックします。

```python
if __name__ == '__main__':
    printer = AutoMailPrinter()
    printer.run()
```

## フォルダ構成
以下のフォルダが存在することを確認してください：
- `path_to_logs`: ログファイルが保存される場所
- `path_to_fax_folder`: エラーメッセージが保存される場所
- `path_to_error_logs`: エラーログが保存される場所
- `path_to_data_files`: 顧客の住所とメール後のデータが保存される場所
- `path_to_output_files`: メール本文がテキストおよびWord形式で保存される場所
- `path_to_pdf_files`: 変換後のPDFファイルが保存される場所
- `path_to_received_fax`: 受信したFAXが保存される場所

## 注意点
- プロジェクトではGmail APIではなくIMAPを使用しています。セキュリティを確保するために、環境変数に認証情報を設定してください。
- ファイルを扱う際、誤って削除されないように重要なファイルのバックアップをとっておくことを推奨します。

## ライセンス
このプロジェクトはMITライセンスのもとで公開されています。
