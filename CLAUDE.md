# CLAUDE.md — text-organizer-studio

「文字整理編輯器」——單檔前端工具：把貼上或上傳的雜亂草稿，透過 AI（固定內建「資深文件排版編輯與文字精煉師」人設）整理成帶一、二、三級標題階層、段落精煉、錯字/標點已校對的正式文件，使用者可在瀏覽器內直接編輯整理後內容，最終下載成 .docx。無建置步驟、無框架、無 package.json，直接開啟 `index.html`（`file://`）或以靜態伺服器託管即可。

## 架構

單一 `index.html`：CSS/JS 全內嵌、無外部資源常駐（`docx@9.7.1` 僅在按下「下載 Word」時才 lazy load）。視覺主題深藍底＋金色 `--accent #e0b64c`（呼應「正式文件／公文」調性），與姊妹工具區隔。

- **輸入區**：`#rawInput` textarea 可直接貼上，或按「上傳 .txt / .docx」讀入既有檔案。`.txt` 走標準 `FileReader.readAsText`；`.docx` 走零依賴的 ZIP 手動解析（`extractDocxText()`/`docxXmlToText()`，直接複製自 `政府補助認證產生器/phoenix-loan-generator/index.html` 已驗證過的實作，不依賴 mammoth.js 等第三方函式庫）。
- **AI 人設固定內建**：`SYSTEM_PROMPT` 常數把使用者提供的【能力設定】【角色定位】【背景洞悉】【任務表述】【語氣人格】【規則實驗】原樣組成，**不開放使用者調整**。額外要求 AI 用 Markdown 標題語法（`#`/`##`/`###`）＋段落間空行輸出全文，並在文末用固定分隔標記「===調整說明===」把「整理後全文」與「調整說明」分開，前端靠這個標記字串比對拆解成兩段（`applyAiResult()`），不依賴 AI 回傳 JSON。
- **「額外需求」欄位**：5 個預設提示詞快速鍵（`EXTRA_PRESETS`）＋自由文字欄位。**刻意鎖定為需要有效 API 金鑰才能使用**（`updateKeyGatedUI()`，依 `hasUsableKey()` 判斷）——因為這個客製化指示只有走 AI 路徑才有意義，陽春規則式備援不會讀取此欄位，欄位與快速鍵在沒有金鑰時整組停用並顯示鎖定提示。
- **BYOK `callLLM()`**：與 `ai-prompt-generator`/`scamper-thinking-generator` 同一套邏輯逐字沿用（Claude/OpenAI/Gemini/OpenRouter 四家、逾時 180 秒、429/500/503/529 自動重試最多 3 次）。設定存 `localStorage`（key `textOrganizerApiConfig`）。
- **陽春規則式備援（無金鑰時）**：`ruleBasedOrganize()` 只做機械式整理——合併同一段落內被誤斷的換行、收斂多餘空白行——**不產生標題階層、不改寫文字**，並在 UI 上明確告知使用者這個限制（`#modeBadge` 即時反映目前是「🤖 AI 整理模式」還是「⚙️ 陽春規則式模式」）。
- **輸出區（可編輯）**：`#outputEditor` 是 `contenteditable` div，AI 回應的 Markdown 先經 `markdownToHtml()`（只處理 `#`/`##`/`###` 標題、`-`/`*` 清單、空行分段的段落，非完整 Markdown 規格）轉成真實 `<h1>/<h2>/<h3>/<p>/<ul>` 標籤。上方工具列用 `document.execCommand('formatBlock'|'bold'|'italic', ...)` 讓使用者手動調整標題階層與粗體/斜體——這就是「透過編輯器整理」的核心互動。視覺上刻意用白色 A4 質感卡片（`.doc-editor`，深色系統 UI 中唯一的亮色區塊）模擬 Word 文件觀感。
- **Word (.docx) 匯出**：`docxItemsFromEditor()` 是 DOM-walker 手法，改寫自 `phoenix-loan-generator` 的 `_docxItemsFromContainer()`——**與原版的差異**：原版把 `h1`/`h2` 都對應到同一個 `HEADING_1`（因為原本只用來輸出單一文件標題），本工具需要三個獨立階層，因此改成 `h1→HEADING_1`／`h2→HEADING_2`／`h3→HEADING_3` 各自獨立對應，且不強制置中。`_docxLinesFromNode()`/`_docxRunsFromLine()`/`_docxParagraphsFromNode()` 三個輔助函式（處理粗體/斜體 run 拆分）逐字複製未改動。下載機制固定：`docx.Packer.toBlob()` → `URL.createObjectURL` → 隱藏 `<a download>` 點擊 → 5 秒後 `revokeObjectURL`。已用 Playwright 風格的瀏覽器自動化端對端驗證：H1 標題正確對應 Word 的 `Heading1` 樣式，其餘段落為一般內文。
- **調整說明**：`#changeSummaryBox` 獨立顯示在輸出區下方，**不計入 Word 匯出內容**（`docxItemsFromEditor()` 只讀取 `#outputEditor` 的子節點）。

## 本次刻意未做（部署上線前留待使用者決定）

比照 `pref-confirm-before-deploy-new-experimental-tool` 記憶——新工具完成後不主動推公開 repo／部署 Pages：

- 頂部跑馬燈、`manual.html` 操作手冊、PWA 加入主畫面、訪客計數器——這些是工作區「已部署上線工具」的標準配件，本工具目前只在本機交付，等使用者決定要上線時再依既有慣例補上（可參照 `scamper-thinking-generator`/`ai-prompt-generator` 的既有做法）
- 序號授權（`member-license-gate`）——使用者本次明確表示先不套用
- 桌面版 exe 打包
- 公開 GitHub repo／GitHub Pages 部署

## 已知限制

- `markdownToHtml()` 只支援 `#`/`##`/`###` 標題、簡單清單與段落，不支援表格、巢狀清單、連結等完整 Markdown 語法——若 AI 回應中夾雜這些語法會被當成一般段落文字（含井字號本身）輸出，非本工具的核心使用情境，暫不處理。
- `.docx` 上傳解析是線性化純文字（比照 `phoenix-loan-generator` 的既有實作），不保留原始格式、表格會轉成 ` | ` 分隔的一行文字，僅供「把既有文件內容帶入重新整理」使用，不是完整還原排版。
- AI 整理模式尚未用真實 API 金鑰端對端驗證（環境內無可用金鑰），僅驗證過陽春規則式備援、標題工具列、Word 匯出、.docx 上傳解析四條路徑；`callLLM()` 邏輯逐字沿用已在其他工具驗證過的實作。

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
