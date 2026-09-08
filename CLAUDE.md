# CLAUDE.md — text-organizer-studio

「文字整理編輯器」——單檔前端工具：把貼上或上傳的雜亂草稿，透過 AI（固定內建「資深文件排版編輯與文字精煉師」人設）整理成帶一、二、三級標題階層、段落精煉、錯字/標點已校對的正式文件，使用者可在瀏覽器內直接編輯整理後內容，最終下載成 .docx。無建置步驟、無框架、無 package.json，直接開啟 `index.html`（`file://`）或以靜態伺服器託管即可。

## 架構

單一 `index.html`：CSS/JS 全內嵌、無外部資源常駐（`docx@9.7.1` 僅在按下「下載 Word」時才 lazy load）。視覺主題深藍底＋金色 `--accent #e0b64c`（呼應「正式文件／公文」調性），與姊妹工具區隔。

- **輸入區**：`#rawInput` textarea 可直接貼上，或按「上傳 .txt / .docx」讀入既有檔案。`.txt` 走標準 `FileReader.readAsText`；`.docx` 走零依賴的 ZIP 手動解析＋`DOMParser` 結構化解析（`extractDocxText()`/`docxXmlToText()`，見下方「表格支援」）。
- **AI 人設固定內建**：`SYSTEM_PROMPT` 常數把使用者提供的【能力設定】【角色定位】【背景洞悉】【任務表述】【語氣人格】【規則實驗】原樣組成，**不開放使用者調整**。額外要求 AI 用 Markdown 標題語法（`#`/`##`/`###`）＋段落間空行輸出全文、遇表格務必用 Markdown 表格語法保留（不可拆散成段落或條列），並在文末用固定分隔標記「===調整說明===」把「整理後全文」與「調整說明」分開，前端靠這個標記字串比對拆解成兩段（`applyAiResult()`），不依賴 AI 回傳 JSON。
- **「額外需求」欄位**：5 個預設提示詞快速鍵（`EXTRA_PRESETS`）＋自由文字欄位。**刻意鎖定為需要有效 API 金鑰才能使用**（`updateKeyGatedUI()`，依 `hasUsableKey()` 判斷）——因為這個客製化指示只有走 AI 路徑才有意義，陽春規則式備援不會讀取此欄位，欄位與快速鍵在沒有金鑰時整組停用並顯示鎖定提示。
- **BYOK `callLLM()`**：與 `ai-prompt-generator`/`scamper-thinking-generator` 同一套邏輯逐字沿用（Claude/OpenAI/Gemini/OpenRouter 四家、逾時 180 秒、429/500/503/529 自動重試最多 3 次）。設定存 `localStorage`（key `textOrganizerApiConfig`）。
- **陽春規則式備援（無金鑰時）**：`ruleBasedOrganize()` 以空行分段的區塊為單位處理——一般段落合併同一段落內被誤斷的換行、收斂多餘空白行；**若該區塊是 Markdown 表格則整段原樣保留**（見下方「表格支援」），不會被強制合併成一行。**不產生標題階層、不改寫文字**，並在 UI 上明確告知使用者這個限制（`#modeBadge` 即時反映目前是「🤖 AI 整理模式」還是「⚙️ 陽春規則式模式」）。
- **輸出區（可編輯）**：`#outputEditor` 是 `contenteditable` div，內容先經 `markdownToHtml()`（處理 `#`/`##`/`###` 標題、Markdown 表格、`-`/`*` 清單、空行分段的段落，非完整 Markdown 規格）轉成真實 `<h1>/<h2>/<h3>/<table>/<p>/<ul>` 標籤——AI 路徑與陽春規則式備援路徑**共用同一套表格判斷/渲染函式**（`isTableBlock()`/`tableBlockToHtml()`），行為一致。上方工具列用 `document.execCommand('formatBlock'|'bold'|'italic', ...)` 讓使用者手動調整標題階層與粗體/斜體——這就是「透過編輯器整理」的核心互動。視覺上刻意用白色 A4 質感卡片（`.doc-editor`，深色系統 UI 中唯一的亮色區塊）模擬 Word 文件觀感。
- **Word (.docx) 匯出**：`docxItemsFromEditor()` 是 DOM-walker 手法，改寫自 `phoenix-loan-generator` 的 `_docxItemsFromContainer()`——**與原版的差異**：原版把 `h1`/`h2` 都對應到同一個 `HEADING_1`（因為原本只用來輸出單一文件標題），本工具需要三個獨立階層，因此改成 `h1→HEADING_1`／`h2→HEADING_2`／`h3→HEADING_3` 各自獨立對應，且不強制置中；另外加上 `table` 標籤支援（`_docxTableFromElement()`，`<th>` 加淺灰底色）。`_docxLinesFromNode()`/`_docxRunsFromLine()`/`_docxParagraphsFromNode()` 三個輔助函式（處理粗體/斜體 run 拆分）逐字複製未改動。下載機制固定：`docx.Packer.toBlob()` → `URL.createObjectURL` → 隱藏 `<a download>` 點擊 → 5 秒後 `revokeObjectURL`。已端對端驗證：H1 標題正確對應 Word 的 `Heading1` 樣式，表格正確匯出成真正的 Word 表格（非純文字），其餘段落為一般內文。
- **調整說明**：`#changeSummaryBox` 獨立顯示在輸出區下方，**不計入 Word 匯出內容**（`docxItemsFromEditor()` 只讀取 `#outputEditor` 的子節點）。

## 表格支援（2026-09-08 修正：原本上傳的 .docx 若含表格，表格會不見）

**根因**：最初的 `docxXmlToText()` 是純 regex 對整份 XML 字串做替換（`</w:tc>` → `" | "`），沒有依 `<w:tr>` 換行，導致整張表格的所有列被壓扁成同一行 pipe 分隔文字；後續 `ruleBasedOrganize()` 又會把同一段落區塊的所有行 `join(' ')`，進一步把表格結構打散成一段看不出列的純文字。

**修法**：
- `extractDocxText()` 改用瀏覽器原生 `DOMParser` 解析 `word/document.xml`（`docxXmlToText()`），依 `w:body` 的直接子節點（`w:p`／`w:tbl`）依序輸出：段落照舊輸出純文字，表格則用 `markdownTableFromTblEl()` 逐列組成標準 Markdown 表格語法（`| A | B |` \n `| --- | --- |` \n `| 1 | 2 |`），儲存格內的 `|` 字元會轉義成 `\|` 避免破壞表格語法。若 `DOMParser` 解析失敗（極少數不規則 XML），退回舊版 regex 邏輯（重新命名為 `docxXmlToTextLegacy()`，僅供保底，仍會有表格被壓成一行的限制）。
- `ruleBasedOrganize()`／`markdownToHtml()` 皆改用共用的 `isTableBlock()`/`tableBlockToHtml()`／`splitTableRow()`/`isTableSeparatorRow()` 判斷「這個以空行分隔的區塊是不是一張 Markdown 表格」（第一行以 `|` 開頭、第二行是 `---` 分隔列），是的話原樣保留（陽春模式）或渲染成 `<table>`（兩條路徑皆適用），不會被當成一般段落合併成一行。
- `SYSTEM_PROMPT` 新增規則，明確要求 AI 遇到表格資料時務必輸出 Markdown 表格語法，不可拆散成段落或條列。
- `docxItemsFromEditor()` 新增 `table` 標籤處理，讓輸出區的 `<table>` 也能正確匯出成 Word 表格（見上方「架構」）。
- 已用實際測試檔（python-docx 產生的 3×3 表格＋前後段落）端對端驗證：上傳後 `#rawInput` 正確顯示 Markdown 表格 → 陽春模式整理後渲染成真正的 `<table>` → 下載 Word 後用 python-docx 讀回，`document.tables` 確認為真正的 Word 表格（非純文字），內容逐格正確。

## 部署

已推公開 GitHub repo `M255525/text-organizer-studio`，用 `.github/workflows/deploy-pages.yml`（比照 `scamper-thinking-generator` 逐字複製）以 Actions workflow 部署 GitHub Pages（非 legacy branch-source，`gh api repos/M255525/text-organizer-studio/pages -f build_type=workflow` 開啟），已上線：<https://m255525.github.io/text-organizer-studio/>。

## 本次刻意未做

- 頂部跑馬燈、`manual.html` 操作手冊、PWA 加入主畫面、訪客計數器——這些是工作區「已部署上線工具」的標準配件，使用者目前只要求「push github page」，尚未要求補齊這些配件；之後若要加，可參照 `scamper-thinking-generator`/`ai-prompt-generator` 的既有做法
- 序號授權（`member-license-gate`）——使用者本次明確表示先不套用
- 桌面版 exe 打包

## 已知限制

- `markdownToHtml()` 支援 `#`/`##`/`###` 標題、Markdown 表格、簡單清單與段落，但不支援巢狀清單、連結、儲存格內格式（粗體/斜體）等完整 Markdown 語法——若 AI 回應中夾雜這些語法會被當成一般文字輸出，非本工具的核心使用情境，暫不處理。
- `.docx` 表格上傳解析目前只取每格的純文字（多段落用空白合併），不保留原始儲存格內的粗體/斜體/合併儲存格（colspan/rowspan）等格式；巢狀表格（表格中的表格）不支援，會被忽略。僅供「把既有文件內容帶入重新整理」使用，不是完整還原排版。
- AI 整理模式尚未用真實 API 金鑰端對端驗證（環境內無可用金鑰），僅驗證過陽春規則式備援、標題工具列、Word 匯出、.docx 上傳解析（含表格）四條路徑；`callLLM()` 邏輯逐字沿用已在其他工具驗證過的實作，AI 是否確實遵守「表格用 Markdown 語法保留」的新規則仍待實測。

## Port

**8812**（工作區 8765-8811 已被其他專案占用，8811 是 `phoenix-coaching-record-generator`）。已在 `.claude/launch.json` 新增對應設定。

## 指令

無建置/測試指令。修改 `index.html` 後直接用瀏覽器開啟驗證，或暫起 `python -m http.server 8812` 測完關閉。修改內嵌 `<script>` 後可用以下方式快速檢查語法：

```bash
python -c "
import re
html = open('index.html', encoding='utf-8').read()
open('_check.js','w',encoding='utf-8').write('\n\n'.join(re.findall(r'<script>(.*?)</script>', html, re.S)))
"
node --check _check.js
```
