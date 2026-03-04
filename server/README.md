# Input Studio Web Server

Web版のInput Studioサーバーです。デスクトップ版と同等の機能をWebブラウザ上で提供します。

## セットアップ

### Windows環境でのセットアップ

**方法1: バッチファイルを使用（推奨）**

```cmd
# セットアップスクリプトを実行
server\setup_windows.bat

# サーバーを起動
server\start_windows.bat
```

**方法2: 手動セットアップ**

```cmd
# 仮想環境を作成
python -m venv .venv

# 仮想環境を有効化
.venv\Scripts\activate.bat

# 依存パッケージをインストール
python -m pip install --upgrade pip
python -m pip install -r server\requirements.txt

# サーバーを起動
python server\main.py
```

**PowerShellの場合:**

```powershell
# 仮想環境を作成
python -m venv .venv

# 仮想環境を有効化
.venv\Scripts\Activate.ps1

# 依存パッケージをインストール
python -m pip install --upgrade pip
python -m pip install -r server\requirements.txt

# サーバーを起動
python server\main.py
```

### Linux/Mac環境でのセットアップ

```bash
# 仮想環境を作成
python -m venv .venv
source .venv/bin/activate

# 依存パッケージをインストール
pip install -r server/requirements.txt

# サーバーを起動
python server/main.py
```

サーバーは `http://localhost:8000` で起動します。

### Dockerを使用した起動

```bash
cd server
docker-compose up -d
```

### 本番環境（ConoHa VPS）

1. サーバーにSSH接続
2. Gitリポジトリをクローン
3. 依存パッケージをインストール
4. systemdサービスを作成（下記参照）
5. Nginxでリバースプロキシ設定

## systemdサービス設定

`/etc/systemd/system/inputstudio.service`:

```ini
[Unit]
Description=Input Studio Web Server
After=network.target

[Service]
Type=simple
User=www-data
WorkingDirectory=/path/to/input-studio
Environment="PORT=8000"
Environment="INPUTSTUDIO_LOCAL_DIR=/var/lib/inputstudio/data"
ExecStart=/usr/bin/python3 /path/to/input-studio/server/main.py
Restart=always

[Install]
WantedBy=multi-user.target
```

サービスを有効化:

```bash
sudo systemctl enable inputstudio
sudo systemctl start inputstudio
```

## Nginx設定例

`/etc/nginx/sites-available/inputstudio`:

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

## 環境変数

- `PORT`: サーバーポート（デフォルト: 8000）
- `INPUTSTUDIO_LOCAL_DIR`: データ保存ディレクトリ（デフォルト: システムのLocalAppData）

## APIエンドポイント

- `POST /api/session` - セッション作成
- `POST /api/projects/create` - プロジェクト作成（PDFアップロード）
- `POST /api/projects/{project_id}/load` - プロジェクト読み込み
- `GET /api/projects/{project_id}/preview/{page_index}` - プレビュー取得
- `POST /api/projects/{project_id}/save` - プロジェクト保存
- `POST /api/projects/{project_id}/values` - タグ値設定
- `POST /api/projects/{project_id}/placements` - テキストフィールド追加
- `GET /api/projects/{project_id}/export` - PDFエクスポート

詳細は `server/main.py` を参照してください。
