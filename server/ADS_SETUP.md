# 本番広告設定メモ（ConoHa VPS + Xserver ドメイン）

## 1) 事前準備

- Google AdSense でサイト審査を通す
- AdSense の `ca-pub-...`（Publisher ID）を控える
- 広告ユニットを 4 つ作成する
  - ゲート画面: `gate`
  - 左パネル上: `panel`
  - 左パネル下: `panelBottom`
  - 機能解放モーダル: `unlock`

## 2) サーバー環境変数

FastAPI を起動する環境で以下を設定してください。

```bash
INPUTSTUDIO_ADS_ENABLED=1
INPUTSTUDIO_AD_PROVIDER=adsense
INPUTSTUDIO_ADSENSE_CLIENT=ca-pub-3765496771126537
INPUTSTUDIO_AD_SLOT_GATE=7232552638
INPUTSTUDIO_AD_SLOT_PANEL=9667144285
INPUTSTUDIO_AD_SLOT_PANEL_BOTTOM=4315074822
INPUTSTUDIO_AD_SLOT_UNLOCK=6857002329
INPUTSTUDIO_UNLOCK_AD_SECONDS=3
# Optional: custom ads.txt line (if omitted, generated from INPUTSTUDIO_ADSENSE_CLIENT)
# INPUTSTUDIO_ADS_TXT_LINE=google.com, pub-3765496771126537, DIRECT, f08c47fec0942fa0
```

`INPUTSTUDIO_ADS_ENABLED=0` にすると広告配信を即停止できます。

## 3) 反映確認

- `https://<your-domain>/ad-config.js` を開く
  - `window.__INPUTSTUDIO_AD_CONFIG__` に設定値が出ること
- `https://<your-domain>/ads.txt` を開く
  - `google.com, pub-..., DIRECT, f08c47fec0942fa0` が出ること
- トップ画面とメイン画面左パネルに広告が表示されること
- ZIPオープン / PDF追加時の機能解放モーダル内に広告が表示されること

## 4) Xserver / ConoHa の使い分け（推奨）

- ConoHa VPS:
  - FastAPI + Uvicorn を常駐（実アプリ）
  - Nginx で `443` 終端、`127.0.0.1:8001` へリバースプロキシ
- Xserver:
  - ドメイン管理・DNS運用（必要ならメールも）
  - `A` レコードを ConoHa のグローバルIPへ向ける

## 5) 注意

- 広告ブロッカー有効時は表示されないことがあります（仕様）
- AdSense 側ポリシーにより広告が出るまで時間差が出る場合があります
- クリック誘導文言（「ここをクリックして」等）はポリシー違反になるため禁止

## 6) まとめて反映する場合

HTTPS化と広告本番反映を一気に進める場合は、以下の運用手順を参照してください。

- `server/deploy/PRODUCTION_HTTPS_ADS_ROLLOUT.md`
