# Input Studio

既存PDFの上に「タグ（入力欄）」を配置し、入力値を重ねて完成PDFを出力するデスクトップアプリケーションです。

## 機能

- PDFテンプレート上にテキスト入力欄を配置
- リアルタイムプレビュー（PNG生成）
- 縦書き・横書き対応
- フォントサイズ・色・行間・文字間隔のカスタマイズ
- 複数ページ対応
- 作業者管理機能
- 完成PDFのエクスポート

## 技術スタック

- **バックエンド**: Python / pywebview
- **フロントエンド**: HTML/CSS/JavaScript
- **PDF処理**: pypdf, reportlab, PyMuPDF, pdf2image
- **画像処理**: Pillow

## セットアップ

### 必要な環境

- Python 3.8以上
- .NET Desktop Runtime（Windowsの場合）

### インストール

```bash
# リポジトリをクローン
git clone https://github.com/TomohikoSASANO/input-studio.git
cd input-studio

# 仮想環境を作成（推奨）
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # Linux/Mac

# 依存パッケージをインストール
pip install -r requirements.txt
```

## 実行（開発環境）

```bash
python app.py
```

## ビルド（PyInstaller）

Windows用の実行ファイルをビルドする場合：

```bash
python -m PyInstaller InputStudio.spec
```

ビルド結果は `dist/InputStudio/` に出力されます。

## 使い方

1. プロジェクトを開く（または新規作成）
2. PDFテンプレートを読み込む
3. テキスト入力欄（タグ）を追加・配置
4. 入力値を設定
5. プレビューで確認
6. 完成PDFをエクスポート

## 注意事項

- OCR機能はありません（フォーム付きPDF前提でもありません）
- プレビューはPNGを生成して表示します
- 日本語フォント（Noto Sans JP）が同梱されています

## ライセンス

[ライセンス情報を追加してください]

## 貢献

プルリクエストやイシューの報告を歓迎します。

## 作者

TomohikoSASANO

