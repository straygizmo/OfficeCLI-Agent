# officecli-agent

[English](README.md) | 日本語

Word / Excel / PowerPoint（Windows デスクトップ版）に組み込む、オープンソースの AI エージェント **VSTO アドイン**です。

- **エージェント**: [GitHub Copilot SDK](https://github.com/github/copilot-sdk)（Copilot CLI のエージェントランタイム）。Copilot のサブスクリプション（CLI ログイン）か、**BYOK**（OpenAI / Azure OpenAI / Anthropic / Ollama などの自前キー）のどちらでも動きます。OAuth の独自実装はありません。
- **文書編集**: [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) をホスト側のツールとして接続（ホストが CLI を直接実行）。ピボット、変更履歴、アニメーション、数式、Mermaid 図、テンプレート差し込み、問題診断、スクリーンショットなど OOXML レベルの編集を高精度に行えます。
- **ライブ編集**: Office の COM オブジェクトモデルで「いま開いている文書」を直接読み書き（Undo 可）。
- **他ファイルの取り込み**: PC 上の `.docx` / `.xlsx` / `.pptx` をエージェントが読み、開いている文書へ差し込みます（Word: 本文/選択範囲へ挿入、PowerPoint: スライド挿入、Excel: シート追加）。

タスクペインとリボンの表示言語は Office の UI 言語に合わせて切り替わります（日本語環境では日本語、それ以外は英語）。

## 仕組み

```
┌─ Office (Word / Excel / PowerPoint) ─────────────────────────────────┐
│  VSTO アドイン (vsto/)                                                │
│   ├ リボン: ホームタブ「エージェント」トグル                           │
│   ├ タスクペイン: WebView2 に React チャット UI (packages/ui) を表示    │
│   ├ HostRelay: UI ⇄ ホストの WebSocket 中継                            │
│   └ *Ops.cs: office.request を COM で実行 (取得/書込/エクスポート/取込) │
└───────────────────────────┬──────────────────────────────────────────┘
                            │ ws://127.0.0.1:4320 (自動起動)
┌─ ローカルホスト (packages/host, Node) ─────────────────────────────────┐
│  GitHub Copilot SDK → Copilot CLI (エージェントループ, Copilot または BYOK) │
│   ├ カスタムツール: office_get_document_context / office_export_document │
│   │                 / office_import_file / word_* / excel_* / ppt_*     │
│   ├ officecli ツール: CLI を直接実行 (読み取りは自動承認、書き込みは確認) │
│   └ 組み込み: ファイル閲覧/検索/編集, シェル (タスクペインで承認)         │
└──────────────────────────────────────────────────────────────────────┘
```

OfficeCLI はディスク上のファイルしか扱えないため、「開いている文書」はアドインがコピーを書き出し、OfficeCLI で編集し、アドインが取り込み直す（ラウンドトリップ）ことでライブに反映します。

| やりたいこと | エージェントが使う経路 |
|---|---|
| セルの書き込み、文章の挿入・置換、スライドのテキスト | `excel_set_range` / `word_insert_text` / `ppt_set_shape_text` など（COM、即時反映・Undo 可） |
| 他の Office ファイルの内容を取り込む | `officecli view/get` で確認 → `office_import_file` |
| ピボット・変更履歴・アニメーション・数式・図など | `office_export_document` → `officecli` で編集 → `office_import_file` |
| 文書の問題診断・見た目の確認 | `office_export_document` → `officecli view … issues / screenshot` |

Excel だけは既存シートを差し替えず「シート追加」または「officecli で計算した結果を `excel_set_range` でセルに適用」に限定し、参照の破壊と Undo 不能を避けています。

## 前提 (Windows)

- Windows 版 Microsoft 365 / Office 2021 以降
- Visual Studio 2022 + 「Office/SharePoint 開発」ワークロード（VSTO アドインのビルド用）
- Microsoft Edge WebView2 ランタイム
- Node.js 20.19 以上、pnpm 9
- [GitHub Copilot CLI](https://github.com/github/copilot-cli)（`copilot` が PATH にある）。Copilot のモデルを使う場合は一度 `copilot` を起動してサインイン。**BYOK だけなら不要**
- [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI)（`officecli` が PATH にある。`irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex`）

## セットアップ

```powershell
pnpm install
pnpm build          # protocol → host → ui
```

次に `vsto\OfficeCliAgent.<App>\OfficeCliAgent.<App>.sln` を Visual Studio で開いて F5 で起動します（詳細は [vsto/README.md](vsto/README.md)）。ホスト（Node）はアドインが自動起動するので、別途立ち上げる必要はありません。

ホストだけを手動で動かしてブラウザからチャット UI を試すこともできます。

```powershell
pnpm start          # http://localhost:4320/taskpane.html?hostApp=Word （Office 操作は失敗します）
```

URL に `&lang=ja` / `&lang=en` を付けると表示言語を固定できます（省略時はブラウザの言語）。

## 設定

### ホスト: `%USERPROFILE%\.officecli-agent\config.json`（初回起動時に生成）

```json
{
  "port": 4320,
  "defaultProvider": "copilot",
  "model": "claude-sonnet-4.5",
  "providers": [],
  "allowedDirectories": ["C:\\Users\\you\\Documents"],
  "officecliPath": "officecli",
  "autoApproveReadOnly": true,
  "permissionTimeoutSeconds": 180
}
```

- `allowedDirectories`: エージェントが確認なしに読み書きできるフォルダ。ここ以外のファイルを `office_import_file` で読む場合はタスクペインに承認プロンプトが出ます。
- `autoApproveReadOnly`: 読み取り専用のツール呼び出し（ファイル閲覧、OfficeCLI の読み取り系コマンド view/get/query/help/validate）を自動承認。シェル実行・ファイル書き換え・URL アクセスは常にタスクペインで承認を求めます。
- `model`: Copilot のモデル ID（省略時は Copilot CLI の既定）。タスクペインの設定からも選べます（新しいセッションから有効）。
- `copilotCliPath`: `copilot` が PATH にない場合の実行ファイルパス。

### アドイン: `%LOCALAPPDATA%\OfficeCliAgent\addin.json`

ポート、`node` のパス、ホスト/UI のフォルダ、自動起動の可否。省略時はアドイン DLL の位置からリポジトリを自動検出します。詳細は [vsto/README.md](vsto/README.md)。

## BYOK（自前の API キー / ローカルモデル）

タスクペインの「設定 → BYOK プロバイダー → 追加」から登録できます（プリセット: OpenAI / Anthropic / Azure OpenAI / Ollama / LM Studio）。登録内容はホストの `config.json` に保存され、API キーはタスクペインには返されません。直接書いてもかまいません。

```json
{
  "defaultProvider": "ollama",
  "providers": [
    { "name": "openai", "type": "openai", "baseUrl": "https://api.openai.com/v1", "apiKeyEnv": "OPENAI_API_KEY", "model": "gpt-6-luna" },
    { "name": "anthropic", "type": "anthropic", "baseUrl": "https://api.anthropic.com", "apiKey": "sk-ant-…", "model": "claude-sonnet-5.5" },
    { "name": "azure", "type": "azure", "baseUrl": "https://<resource>.openai.azure.com", "apiKey": "…", "model": "gpt-6-luna", "azureApiVersion": "2024-10-21" },
    { "name": "ollama", "type": "openai", "baseUrl": "http://localhost:11434/v1", "model": "llama3.1", "modelProvider": "ollama" }
  ]
}
```

- `type`: `openai`（OpenAI 互換全般: OpenAI / Ollama / vLLM / LiteLLM / LM Studio / Foundry Local）、`azure`（Azure OpenAI ネイティブ）、`anthropic`。
- `wireApi`: 省略時は OpenAI（`api.openai.com`）と Azure OpenAI は `responses`（gpt-5.6 以降を前提）、それ以外の OpenAI 互換サーバーは `completions`。必要なら明示できます。Azure の `model` はデプロイ名です。
- `apiKey` の代わりに `apiKeyEnv` で環境変数名を指定できます。ローカルモデルはキー不要。
- BYOK ではモデル ID が必須です（プロバイダーの既定 `model`、またはタスクペインで入力）。
- BYOK 時の利用量・レート制限はそのプロバイダー側のもので、Copilot のプレミアムリクエストは消費しません。ツール呼び出しを扱えるモデルを選んでください（Ollama なら Llama 3.1 以上、Qwen 2.5 系など）。

## 使い方の例

- 「この文書の構成を要約して」
- 「`C:\Users\me\Documents\売上.xlsx` の Sheet1 をこのブックの末尾に追加して」
- 「`提案書_旧.pptx` の 3〜5 枚目をこのプレゼンの 2 枚目の後ろに入れて」
- 「B2:B20 の合計を B21 に入れて、通貨書式にして」
- 「このプレゼンのレイアウト崩れ（テキストのはみ出しなど）を診断して直して」

## リポジトリ構成

```
packages/protocol   アドイン ⇄ ホスト間のメッセージ型 (WebSocket)。OfficeOp の定義もここ
packages/host       ローカルホスト (Node): Copilot SDK セッション、OfficeCLI ツール、Office ツール、BYOK
packages/ui         チャット UI (Vite + React)。WebView2 の postMessage とブラウザ用 WebSocket の両対応
vsto/               VSTO アドイン: 共有ライブラリ (Core) と Word / Excel / PowerPoint の各ソリューション
```

## 設計上の注意: OfficeCLI と起動中の Office

OfficeCLI の `view … screenshot` / `stats` は既定（`--render auto`）で Windows にインストール済みの Word / PowerPoint を COM で操作して描画します。アドインが動作中の Office に外部プロセスから接続すると PowerPoint がクラッシュするため、ホストは OfficeCLI を MCP サーバーではなく自前のツールとして実行し、これらのコマンドには常に `--render html`（内蔵レンダラー）を付けています。シェルから PowerShell 等で Office を COM 操作することもシステムプロンプトで禁止しています。

## セキュリティ上の注意

- エージェントはローカルファイルの読み取りと（Copilot CLI 組み込みの）Web アクセスの両方を持ちます。ファイルや Web から得た文章に埋め込まれた指示は「データ」として扱うようシステムプロンプトで指示していますが、シェル実行・書き込み・URL アクセスは必ずタスクペインで承認してください。
- ホストは `127.0.0.1` にのみバインドし、認証はありません。同一 PC 上の他プロセスからも接続できる点に留意してください。
- Office on the web / Mac / iPad では動作しません（VSTO・ローカルホスト前提）。

## クレジット

このプロジェクトは次のオープンソースの上に成り立っています。

- **[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI)** (iOfficeAI, Apache-2.0) — Word / Excel / PowerPoint ファイルの読み書き・レンダリング・診断を担う文書エンジン。本アドインの「高度な編集機能」はすべて OfficeCLI の CLI をホストから呼び出して実現しています。
- **[GitHub Copilot SDK](https://github.com/github/copilot-sdk)** (GitHub, MIT) — エージェントランタイム。ツール呼び出しループ、MCP 接続、BYOK、権限制御は Copilot SDK / Copilot CLI の機能です。
- [office-agents](https://github.com/hewliyang/office-agents) (hewliyang, MIT) — Office アドインに BYOK エージェントを組み込むという設計の参考にしました。
- [Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/) — タスクペインの UI ホスト。

OfficeCLI と GitHub Copilot SDK はそれぞれ別途インストールが必要で、本リポジトリには同梱していません。各プロジェクトのライセンスと利用規約に従ってください。

## ライセンス

本リポジトリのコードは MIT ライセンスです。
