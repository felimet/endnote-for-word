# EndNote for Word

[English](README.md)

可安裝的 AI Agent skill、Codex/Claude plugin 與 MCP server，依可審閱的 JSON manifest，在 Microsoft Word 文件中建立真正的 EndNote Cite While You Write (CWYW) 引用與參考文獻清單。

## 本工具做什麼

- 將 DOCX 中的 `{{CITE:id}}` 標記換成真正的 CWYW 引用：每一筆引用都是帶有 EndNote traveling library 的 `ADDIN EN.CITE` 欄位，參考文獻清單則是 `ADDIN EN.REFLIST` 欄位。
- 不手打引用文字、不在 OOXML 中偽造欄位碼，也不以純文字清單假冒參考文獻清單。
- 以唯讀方式讀取 `.enl` 書目庫，讓 Agent 在開啟 Word 之前就能搜尋文獻並確認每一個紀錄編號。
- 預設採 IEEE，每筆被引用文獻都必須具備 DOI，引用前會透過 doi.org 驗證，並拒絕重複的 resolver 前綴。
- 同時支援 IEEE Default Format 的 `[x]`，以及 EndNote Display as Author (Year) 所產生的 `Author et al. [x]` 敘述式引用。
- 在渲染前先對 manifest 與已安裝輸出樣式做預先檢查。不存在的紀錄編號或樣式名稱會讓 EndNote 跳出對話框，automation 會卡住直到 process 被強制結束。
- Word 全程隱藏；由 bridge 啟動的 EndNote 也以隱藏視窗執行。
- 以程式方式切換文件輸出樣式並重建參考文獻清單，全程不需對話框。
- 對完成的文件做結構性驗證，回報欄位數量、樣式，以及文件記錄的書目庫資訊。
- 匯出 RIS 檔，讓尚未收錄的文獻可以一次匯入書目庫。

## 系統需求

- Windows 10 或 11。
- 支援 COM automation 的 Microsoft Word 桌面版。
- 桌面版 **EndNote** for Windows，且已安裝並啟用 **Cite While You Write** Word 增益集。此為硬性前提：沒有 CWYW 增益集就沒有可支援的方式產生原生引用欄位，本工具也不會偽造。
- `PATH` 中可使用 Python 3。
- PowerShell 7（`pwsh.exe`）或 Windows PowerShell 5.1（`powershell.exe`）。
- 可透過 HTTPS 存取 `doi.org`，供引用前的強制 DOI 驗證。

本專案以 **EndNote 21** 及其 64 位元 CWYW 增益集（`EndNote.WordAddin.Connect`）與全域 Word 範本 `EndNote Cwyw.dotm` 開發與驗證。EndNote 執行檔、樣式目錄與 CWYW 範本皆於執行期由登錄檔與安裝目錄探索，並以最新版本優先，程式中不寫死版本號。

EndNote 與 Microsoft Office 為專有產品，本 repository 不包含其安裝程式或授權。EndNote **web** 或 **online** 外掛不能取代桌面版 CWYW 增益集。

## 終端機相容性

Bridge 為單一 PowerShell 腳本，同時相容 PowerShell 7 與 Windows PowerShell 5.1，共支援四種終端機。

| 目前終端機 | 執行方式 |
|---|---|
| PowerShell 7+ | 直接以 `pwsh.exe` 執行命令。 |
| Windows PowerShell 5.1 | 支援。可用 `powershell.exe` 執行；若已安裝 `pwsh.exe` 則優先使用。 |
| Windows 上的 Bash，包括 Git Bash | 呼叫 Windows `pwsh.exe`，找不到時退回 `powershell.exe`。 |
| CMD | 以相同參數呼叫 `pwsh.exe`，找不到時退回 `powershell.exe`。 |

在 WSL Bash 中須呼叫 Windows 側的執行檔；Linux `pwsh` 無法操作 Windows Word COM。若 `pwsh.exe` 與 `powershell.exe` 皆不存在，應停止執行並安裝 PowerShell，不可默默降級。安裝說明見 [PowerShell 安裝指南](https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows)。

以下命令可在 PowerShell 7、Windows PowerShell 5.1、Git Bash 與 CMD 執行：

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/endnote-word.ps1 -Action probe
```

## 透過 AI Agent 安裝

若要在 Claude Code、Claude Desktop、Codex 或 ChatGPT Desktop 使用本工具，貼上以下 prompt，Agent 會依目前環境完成設定：

```text
從 https://github.com/felimet/endnote-for-word 安裝或升級 EndNote for Word toolkit。偵測我目前可用的終端機，並使用 PowerShell 7、Windows PowerShell 5.1、Windows 上的 Bash，包括 Git Bash，或 CMD；有 pwsh.exe 時優先使用，否則退回 powershell.exe。在 WSL Bash 中呼叫 Windows 側執行檔，不使用 Linux pwsh。兩者皆不存在時，停止執行並提示我依 https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows 安裝 PowerShell。確認 Microsoft Word 桌面版，以及安裝了 Cite While You Write 增益集的桌面版 EndNote 均可使用；安裝可攜式 skill、註冊本機 stdio MCP server、執行 probe 與 repository 測試、保留既有 Agent 設定，並列出所有異動檔案。probe 未回報 cwyw_macro_callable 為 true，或 validation 未回傳 ok: true 時，不得宣稱完成。
```

再告訴 Agent 要處理的 DOCX 與要引用的 `.enl` 書目庫。若未另外指定，使用 IEEE。

各平台 skill 位置與 MCP 設定見[安裝矩陣](skills/endnote-for-word/references/installation-matrix.md)。

## 將 Skill 與 MCP 加入各 AI Agent

先將 `<REPO_ROOT>` 換成 repository 的本機絕對路徑。修改設定檔時必須保留原有 MCP 項目。

### Codex

將 `skills/endnote-for-word` 複製至 `%USERPROFILE%\.codex\skills\endnote-for-word` 或 `%USERPROFILE%\.agents\skills\endnote-for-word`，再註冊 MCP server：

```console
codex mcp add endnote-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
codex mcp get endnote-for-word
```

Release 亦包含 `.codex-plugin/plugin.json`、`.codex-mcp.json` 與 `dist/endnote-for-word-plugin.zip`，供支援 plugin 的部署流程使用。

### Claude Code

將 `skills/endnote-for-word` 複製至 `%USERPROFILE%\.claude\skills\endnote-for-word`，再註冊 MCP server：

```console
claude mcp add --scope user endnote-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
claude mcp get endnote-for-word
```

Release 亦包含 `.claude-plugin/plugin.json`、`.mcp.json` 與合併 plugin 套件。安裝後重新啟動 Claude Code。

### Claude Desktop

於 **Customize > Skills** 上傳 `dist/endnote-for-word.skill`。再將下列 server 合併至 `%APPDATA%\Claude\claude_desktop_config.json`，不可覆蓋其他既有 server，完成後重新啟動 Claude Desktop：

```json
{
  "mcpServers": {
    "endnote-for-word": {
      "command": "pwsh.exe",
      "args": ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", "<REPO_ROOT>\\scripts\\run-mcp.ps1"]
    }
  }
}
```

參閱 Anthropic 的 [Skill 上傳說明](https://support.claude.com/en/articles/12512180-use-skills-in-claude)與[本機 MCP 設定指南](https://modelcontextprotocol.io/docs/develop/connect-local-servers)。

### ChatGPT Desktop

若帳號或 workspace 提供 **Plugins**，依官方 [ChatGPT plugin 流程](https://help.openai.com/en/articles/20001256)安裝或啟用封裝能力。ChatGPT Desktop 不會自動探索本機 repository。OpenAI 目前的 full MCP 文件適用於 ChatGPT web；除非帳號中確實顯示相同的 developer-mode 介面，不可假設桌面版具備相同能力。

ChatGPT 無法直接連接本專案隨附的本機 stdio MCP 命令；必須建立 remote MCP endpoint 或使用 [Secure MCP Tunnel](https://help.openai.com/en/articles/12584461-developer-mode-and-full-mcp-connectors-in-chatgpt-beta)，並依帳號或 workspace 權限啟用 Developer mode 或 Apps。EndNote 與 Word automation 必須在登入中的互動式 Windows 桌面執行，因此 endpoint 或 tunnel 仍須將工具執行路由回該 Windows 主機。本 repository 目前只提供本機 stdio server。

## 工作流程

1. **Probe。** 確認 Word COM、EndNote 執行檔、CWYW 增益集註冊狀態、已載入的 CWYW 範本，以及 CWYW 巨集可被呼叫。巨集檢查在拋棄式空白文件上執行，不會動到你的檔案。
2. **選定樣式。** 列出已安裝的 `.ens` 輸出樣式並取得精確名稱。無法對應到已安裝檔案的樣式名稱會讓 EndNote 跳出對話框並卡住整個流程。
3. **查找文獻。** 以唯讀方式搜尋 `.enl` 書目庫，取得紀錄編號、第一作者姓氏、年份與裸 DOI。
4. **撰寫 manifest。** 每一筆引用一個項目，並依語句功能選擇 `display`。作者是句子主詞或具名行動者時用 `author-year`；一般論述證據與合併引用用 `default`。
5. **預先檢查。** 驗證每筆紀錄存在、未在垃圾桶、姓氏、年份與 DOI 相符、EndNote DOI 欄位只含裸 DOI，且 DOI 可由 doi.org 成功解析。
6. **靜默 Render。** Bridge 隱藏 Word，必要時隱藏啟動 EndNote，建立原生引用欄位，只對指定引用套用 `ENAuthorYearFormat`，最後以不可分割的方式儲存。
7. **Validate。** 以唯讀方式重新開啟結果，逐項執行下列驗證閘門。

先在 DOCX 中放好 marker，例如：

```text
自動化跛行評分已成為常規做法 {{CITE:thomsen2023}}。
兩組獨立試驗結論一致 {{CITE:merged}}。
```

## Manifest 契約

完整檔案見 [`examples/example-manifest.json`](examples/example-manifest.json)。該檔的 `library.path` 僅為佔位路徑，使用前請改成你自己的 `.enl` 書目庫路徑，否則 manifest 會在 `library.path` 存在性檢查失敗。

```jsonc
{
  "schema_version": 1,
  "library": { "path": "C:\\EndNoteLibraries\\ExampleLibrary.enl" },
  "style": "IEEE",
  "citations": [
    {
      "id": "thomsen2023",
      "marker": "{{CITE:thomsen2023}}",
      "targets": [ { "rec_number": 1, "author": "Thomsen", "year": 2023, "doi": "10.0000/fixture.0001" } ],
      "display": "default",
      "pages": "45-46"
    }
  ],
  "bibliography": { "expected": true }
}
```

| 規則 | 要求 |
|---|---|
| `schema_version` | 必須為 `1`。 |
| `citations` | 非空陣列。未知的最上層鍵與引用層鍵一律拒絕。 |
| `id` | 符合 `^[A-Za-z][A-Za-z0-9_-]*$`，且全域唯一。 |
| `marker` | 全域唯一，且必須在文件中恰好出現一次。 |
| `targets` | 非空。每個 target 需有大於零的整數 `rec_number`、非空的 `author` 姓氏、整數 `year`，以及不含 `doi:` 或 resolver URL 的裸 `doi`。 |
| `display` | IEEE 下，`default` 產生 `[x]`；`author-year` 產生 `Author et al. [x]`。`author-year` 只能有一個 target。 |
| `pages` | 選填。符合 `^[0-9]+(\s*-\s*[0-9]+)?$`。僅在 `targets` 恰有一個項目時允許。 |
| `suffix` | 選填自由文字，例如 `", p. 99"`。僅在 `targets` 恰有一個項目時允許。 |
| `library.path` | 必須存在，且同層需有 `<name>.Data` 目錄。 |
| `style` | 選填的非空 `.ens` 樣式。省略時使用已設定的預設值 `IEEE`。 |

每一筆引用產生一個暫時引用字串，因此含多個 target 的合併引用只會產生一個欄位，而非多個：

```text
{Thomsen, 2023 #1}
{Thomsen, 2023 #1@45-46}
{Magrin, 2023 #2; Bernhard, 2020 #3}
```

`pages` 與 `suffix` 只作用於最後一個 target，因此限定只能用於單一 target 的引用。

### IEEE DOI 輸出樣式

EndNote 的 Bibliography Templates 必須在 `DOI` 欄位前放入固定文字 `doi: https://doi.org/`，而 EndNote DOI 欄位本身只存裸 DOI。Journal Article、Conference Paper 與 Web Page 的範本規格見 [styles-and-formatting.md](skills/endnote-for-word/references/styles-and-formatting.md)。內建樣式需用 **Save As** 儲存至個人 Styles 目錄，不可覆寫安裝目錄中的原檔。

Manifest 預先檢查會拒絕已帶 `doi:` 或 `https://doi.org/` 的 EndNote DOI 欄位。RIS 匯出則會把這些輸入正規化成唯一的裸 `DO` 值。文件驗證最後確認每筆引用都實際顯示 `doi: https://doi.org/<DOI>`，且沒有 `https://doi.org/https://doi.org/...`。

## 驗證閘門

任一閘門失敗時，`validate` 會回傳 `ok: false` 並以離開碼 `3` 結束。

1. 文件中不得殘留任何 `{{CITE:` 標記文字。
2. 文件中不得殘留暫時引用文字，以 `\{[^{}]*#\d+[^{}]*\}` 比對全文。
3. 原生引用欄位數量等於 manifest 的引用筆數。計數規則為：欄位碼以 `ADDIN EN.CITE` 開頭者的數量，減去以 `ADDIN EN.CITE.DATA` 開頭者的數量；traveling library 內容較大時，EndNote 會將其移入巢狀資料欄位。
4. `bibliography.expected` 為 true 時，必須恰有一個 `ADDIN EN.REFLIST` 欄位，且其結果文字非空白。
5. 每筆 DOI 都需顯示為 `doi: https://doi.org/<DOI>`，且 resolver 前綴不可重複。
6. manifest 中每個紀錄編號都要以 `<rec-number>N</rec-number>` 出現在某個引用欄位的 traveling library XML 中，且對應的作者姓氏出現在同一欄位。
7. 每個 `author-year` 原生欄位的顯示結果都包含作者姓氏。
8. 文件文字中不得出現 `Error!`、`Reference source not found`、`{Citation}` 或 `{Bibliography}` 等占位字串。
9. `EN.Layout` 文件變數存在，且其樣式與 manifest 的樣式相符（不分大小寫）。
10. `EN.Libraries` 文件變數存在，且其 `<record-ids>` 中列出 manifest 的每個紀錄編號。

## 第一版限制

| 限制 | 原因 |
|---|---|
| 不支援引用前綴 | 反斜線前綴語法（例如 `{see \Thomsen, 2023 #1}`）會開啟 EndNote 對話框並卡住 automation。manifest 直接拒絕 `prefix`。 |
| 不支援排除作者或年份 | `{\Thomsen, 2023 #1}` 與 `ENExcludeAuthor` 系列巨集皆由對話框驅動，會卡住流程。樣式允許時請改用 `suffix` 文字。 |
| 書目庫為唯讀 | 本工具不會寫入 `.enl` 檔或其 `.Data` 目錄；它會將內部資料庫複製到暫存檔、讀取複本，再刪除複本。 |
| 新增文獻為半自動 | 新文獻會匯出成 RIS 檔，由你自行在 EndNote 以 **File > Import** 匯入。本工具不寫入你的書目庫。 |
| 參考文獻清單位置採 CWYW 預設 | EndNote 會把清單放在 Cite While You Write 預設的位置，通常是文件結尾。需要調整時請在 EndNote 或 Word 中自行搬移，本工具不負責移動位置。 |
| 不切換 instant formatting | CWYW 的切換巨集無法讀回目前狀態，盲目呼叫會翻轉未知狀態。需要調整時請在 EndNote 功能區操作。 |
| 轉為純文字須手動操作 | 對應的 CWYW 巨集 `ENRemoveFieldCodes` 會跳出確認對話框，在 automation 下永不回傳，因此本工具沒有 `remove-field-codes` 動作，也沒有對應的 MCP 工具。請自行在 Word 功能區以 **EndNote > Convert Citations and Bibliography > Convert to Plain Text** 執行；該操作會另存新檔，原始檔的欄位碼保持不變。轉換結果不可逆，EndNote 無法再更新該複本中的引用。 |

## 驗證安裝

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/endnote-word.ps1 -Action probe
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/endnote-word.ps1 -Action list-styles -Style IEEE
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File tests/run-tests.ps1
```

Bridge 會保留來源檔，透過同層暫存 DOCX 完成儲存；除非明確指定 `-Overwrite`，否則不覆寫既有輸出。即時整合測試需自行啟用，並提供真實書目庫路徑。

## 快速測試 Prompt

先在測試 DOCX 中放入唯一標記 `{{CITE:test}}`，再將以下 Prompt 中的三個路徑或 DOI 換成實際值後交給 AI Agent：

```text
請使用 endnote-for-word，將 EndNote library「<完整 .enl 路徑>」中 DOI 為「<裸 DOI>」的文獻插入 Word「<完整 .docx 路徑>」的 {{CITE:test}} 位置。保留來源檔並另存新檔，預設使用 IEEE；若標記是句子的作者主詞，使用 Display As: Author (Year)，否則使用 Default Format。Word 與 EndNote 全程靜默。引用前先驗證 DOI，完成後驗證原生 EndNote citation 與 bibliography fields，並確認 bibliography 顯示 doi: https://doi.org/<DOI>，最後回報輸出路徑與驗證結果。
```

## Repository 結構

| 路徑 | 用途 |
|---|---|
| `skills/endnote-for-word/` | 跨 Agent skill、參考資料及 launcher |
| `scripts/` | Word 與 EndNote automation bridge、書目庫讀取器、MCP server 及封裝程式 |
| `config/defaults.json` | 預設輸出樣式、暫時引用分隔符號與驗證閘門 |
| `examples/` | 範例 manifest |
| `tests/` | 靜態、MCP 協定、離線結構、書目庫及即時整合測試 |
| `evals/` | 評測案例 |

## 封裝

```console
python scripts/package_plugin.py
```

命令會產生 `dist/endnote-for-word-plugin.zip` 及 SHA-256 檔。獨立 skill 套件為 `dist/endnote-for-word.skill`。

## 問題與交流

若有任何問題，請至 GitHub Issues 提出並參與交流。

## 授權

[MIT](LICENSE)
