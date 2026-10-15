# vsto/ — VSTO アドイン (Word / Excel / PowerPoint)

```
OfficeCliAgent.Core/                      共有ライブラリ (net48, SDK スタイル csproj)
  AgentSettings.cs                          設定 (%LOCALAPPDATA%\OfficeCliAgent\addin.json) とログ
  HostProcessManager.cs                     Node ホストの検出・自動起動 (--port/--no-tls/--idle-exit)
  HostRelay.cs                              WebView2 ⇄ WebSocket 中継。office.request を COM で実行
  AgentTaskPane.cs                          WebView2 を載せた UserControl (チャット UI を表示)
  AgentController.cs                        上記をまとめる
  IOfficeOps.cs                             アプリ別 COM 操作のインターフェース + ヘルパー
OfficeCliAgent.Word/OfficeCliAgent.Word.sln         Word VSTO アドイン (ソリューション)
OfficeCliAgent.Word/OfficeCliAgent.Word/            ThisAddIn.cs, Ribbon.cs, Ribbon.xml, WordOps.cs + VS 生成ファイル
OfficeCliAgent.Excel/OfficeCliAgent.Excel.sln       Excel  (ExcelOps.cs)
OfficeCliAgent.PowerPoint/OfficeCliAgent.PowerPoint.sln  PowerPoint (PowerPointOps.cs)
```

各 VSTO プロジェクトは Visual Studio の「<App> VSTO アドイン」テンプレートから生成し、`OfficeCliAgent.Core` をプロジェクト参照、NuGet で `Microsoft.Web.WebView2` と `Newtonsoft.Json` を追加、`Ribbon.xml` を埋め込みリソースにしてあります。`ThisAddIn.Designer.cs` などの生成ファイルはテンプレートのままです。

## 前提

- Visual Studio 2022 で「Office/SharePoint 開発」ワークロードをインストール済み
- .NET Framework 4.8 開発ツール
- Microsoft Edge WebView2 ランタイム（Windows 11 / Microsoft 365 環境なら入っています）
- Node.js 20.19+（`node` が PATH にあること）、GitHub Copilot CLI、OfficeCLI（`officecli` が PATH にあること）
- リポジトリルートで `pnpm install && pnpm build` 済み（`packages/host/dist`, `packages/ui/dist` が必要）

## ビルドと起動

1. 使いたいアプリのソリューション（例: `OfficeCliAgent.Word\OfficeCliAgent.Word.sln`）を Visual Studio で開く。
2. 初回は NuGet パッケージの復元が走る（`packages/` フォルダは git 管理外）。
3. 初回のみ、マニフェスト署名用のテスト証明書を作る。署名用の `*_TemporaryKey.pfx` は秘密鍵なのでリポジトリに含めていません（`.gitignore` で除外）。そのままビルドすると証明書が見つからないエラーになるので、プロジェクトのプロパティ →「署名」→「テスト証明書の作成...」で作り直してください（パスワードは空で可）。`.csproj` の `ManifestKeyFile` / `ManifestCertificateThumbprint` が書き換わりますが、これは各自の環境用の変更なのでコミット不要です。
4. F5 で Office が起動し、タスクペインが自動で開く。ホーム タブ右端の「OfficeCLI Agent」グループのトグルで開閉できる。

新しい VSTO プロジェクトを足す場合（別アプリや作り直し）は、テンプレートで生成 → `ThisAddIn.cs` を上書き → `Ribbon.cs` / `Ribbon.xml`（埋め込みリソース）/ `<App>Ops.cs` を追加 → Core を参照 → NuGet 追加、の順です。

## 動作の流れ

1. ボタンでタスクペインを開く → `AgentTaskPane` が WebView2 を初期化し、`packages/ui/dist` を `https://officecli-agent.local/` にマップして表示。
2. `HostProcessManager` が `127.0.0.1:4320` を確認。何も居なければ `node packages/host/dist/cli.js --port 4320 --no-tls --idle-exit 900` を非表示で起動。
3. `HostRelay` が `ws://127.0.0.1:4320/ws` に接続。ページの `hello` に実際のアプリ名・文書名を書き込んで転送。
4. エージェントのツール呼び出しのうち `office.request` だけを横取りし、Office の UI スレッドで COM 実行して `office.result` を返す。それ以外のフレームはそのままページへ流す。
5. ホストは全アドインが切断して 15 分経つと自動終了（複数の Office アプリで 1 プロセスを共有）。

## 設定

`%LOCALAPPDATA%\OfficeCliAgent\addin.json`（初回起動時に生成）

```json
{
  "Port": 4320,
  "NodePath": "node",
  "HostDir": "C:\\Users\\you\\source\\repos\\officecli-agent\\packages\\host",
  "UiDir": "C:\\Users\\you\\source\\repos\\officecli-agent\\packages\\ui\\dist",
  "AutoStartHost": true,
  "HostIdleExitSeconds": 900
}
```

`HostDir` / `UiDir` を省略すると、アドイン DLL の場所から上位へ辿ってリポジトリ（`packages/host/dist/cli.js`）を探します。VSTO の出力は `vsto\OfficeCliAgent.Word\bin\Debug\` なので、リポジトリ内で開発している限り自動で見つかります。

ログ: `%LOCALAPPDATA%\OfficeCliAgent\addin.log`（ホストの標準出力も `host>` として記録）。

## COM 操作の対応表

| プロトコル `OfficeOp` | Word | Excel | PowerPoint |
|---|---|---|---|
| `getContext` | 段落数・見出し・選択・本文プレビュー | シート一覧・使用範囲・先頭 40 行 | スライド一覧（図形名とテキスト）・レイアウト一覧 |
| `exportFile` | `Content.ExportFragment(.docx)` | `Workbook.SaveCopyAs` | `Presentation.SaveCopyAs` |
| `importFile` | `Range.InsertFile`（Start/End/Replace/Selection） | 非表示で開いて `Worksheet.Copy(After:=)` | `Slides.InsertFromFile`（範囲・位置指定） |
| アプリ固有 | getSelection / insertText(+style) / replaceText | listSheets / getRange / setRange(values, formulas, numberFormat) | setShapeText / addSlide(layout, title, body) / deleteSlides |

## 既知の制約・今後

- Word の `exportFile` は本文範囲のエクスポートです。文書プロパティや一部の設定は複製されません。
- PowerPoint の `importFile` は常に挿入先テーマを適用します（`keepSourceFormatting` は無視して結果に明記）。
- 汎用スクリプト実行（Office.js の代替として Roslyn で C# を実行する `office_run_csharp`）は未実装です。必要になったら `Microsoft.CodeAnalysis.CSharp.Scripting` を Core に追加して `IOfficeOps` に 1 ケース足すだけで入ります。
- 配布は ClickOnce（VSTO 標準）を想定。Node ホストと UI の同梱は未対応で、現状はリポジトリのビルド成果物を参照します。
