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

## 圖片／影片保留（2026-09-15 新增：上傳 .docx 若含圖片或影片，整理與下載時一併保留）

**動機**：使用者要求「上傳要求整理檔案中如果有圖片或影片請一併保留」，且「整理下載時也一併保留」——原本 `extractDocxText()` 只解析 `word/document.xml` 的純文字，圖片／影片會直接消失。

**做法（佔位標記貫穿整條管線）**：因為 AI／陽春規則式整理兩條路徑都只處理純文字，圖片/影片無法直接塞進 `#rawInput` textarea，所以改用「文字佔位標記」貫穿整條管線，最後再換回真正的媒體元素：

- **擷取（`extractDocxText()`）**：改為先列出整份 docx zip 的所有檔案條目（`listDocxZipEntries()`/`readZipEntryBytes()`，取代原本只找單一 `word/document.xml` entry 的邏輯），額外讀取 `word/_rels/document.xml.rels` 建立 r:id → 媒體檔案相對路徑對照表，並預先讀出 `word/media/` 下副檔名符合圖片／影片的位元組。`docxXmlToText()` 在走訪段落時，遇到 `<w:drawing>` 元素會呼叫 `tokenForDrawing()`：先找子孫元素 `<a:videoFile r:link="rIdX">`（Word「插入本機視訊」時，圖片是縮圖、這個元素的 `r:link` 才指向實際影片檔）判斷是否為影片，否則找 `<a:blip r:embed="rIdX">` 判斷為一般圖片；解析出的位元組分別轉成 `data:` URL（圖片，存進 `images` map）或用 `URL.createObjectURL` 的 blob URL（影片，較省記憶體，存進 `videos` map），並在原文字位置插入 `[[IMG_n]]`／`[[VID_n]]` 純文字佔位標記，取代原本的圖片/影片內容。`extractDocxText()` 回傳值從純文字字串改成 `{text, images, videos}`，寫入模組層級變數 `uploadedMedia`（上傳新檔案或按「清空重來」時會呼叫 `revokeUploadedVideos()` 釋放舊的 blob URL 再重置）。
- **整理（AI／陽春規則式皆不變邏輯，只加尾端還原）**：`[[IMG_n]]`／`[[VID_n]]` 是純 ASCII 文字，會像一般文字一樣原封不動撐過 `ruleBasedOrganize()`（陽春模式的行合併）與 AI 整理（`SYSTEM_PROMPT` 新增規則 5，明確要求 AI 原封不動保留這些標記、不要翻譯/刪除/搬移）。`applyAiResult()`／`runFallback()` 產生 HTML 後，統一多呼叫一層 `restoreMediaPlaceholders(html)`，把標記字串換成真正的 `<img class="doc-img">` 或 `<span class="doc-video-wrap">`（內含 `<video controls>` 預覽＋`<a download>` 下載原始影片檔按鈕＋一行提示文字），再寫入 `#outputEditor`。
- **Word 匯出（`docxItemsFromEditor()`／`_docxLinesFromNode()`）**：DOM-walker 新增兩種節點類型——`<img>` 標籤轉成一筆 `{img:true, el}` marker，`_docxRunsFromLine()` 遇到就呼叫 `_docxImageRunFromImgEl()`（讀 `naturalWidth/Height`，超過 600px 寬則等比縮小，`data:` URL 轉 `Uint8Array` 後交給 `docx.ImageRun`，**必須帶 `type` 欄位**——`docx@9.7.1` 的 `ImageRun` 若省略 `type` 只給 `data`，匯出的 docx 會在 `[Content_Types].xml` 缺對應內容類型，Word/python-docx 開啟時報錯，這是實測踩到的坑，`_docxImageTypeFromDataUrl()` 依 `data:image/xxx` 的 MIME 對應到 docx.js 支援的 `png`/`jpg`/`gif`/`bmp`/`svg`）；`.doc-video-wrap` 節點（用 class 判斷，避免落入通用 `else{walk(child)}` 分支把裡面的下載連結文字也重複輸出一次）不遞迴，改成輸出一行純文字提示（含檔名），**影片本身不會匯出進 Word 檔**——因為 Word 文件格式本身就不支援內嵌可播放的本機影片，`docx.js` 也沒有對應 API，這是格式層級的硬限制，只能靠編輯區內的「下載原始影片檔」按鈕讓使用者另外保存原始檔。
- **UI 提示**：輸入區新增一行 hint 說明佔位標記機制；`.warn-box` 新增一條警語明確告知「圖片會保留並可下載回 Word，影片受格式限制無法內嵌播放、下載 Word 時只會留檔名提示，請另外下載影片原始檔」。

**已知限制**：
- 只處理 `word/media/` 底下、副檔名為常見圖片（png/jpg/gif/bmp/webp/svg/tiff）或影片（mp4/mov/avi/wmv/webm/mkv/m4v）格式的檔案；webp/tiff 這類 docx.js 不支援匯出的圖片格式，`_docxImageTypeFromDataUrl()` 會 fallback 標成 `png` 但實際位元組仍是原格式，Word 開啟該圖仍可能失敗——這兩種格式本來就極少出現在 Word 文件內嵌圖片中，暫不特別處理。
- 影片位置定位僅支援「本機插入視訊」（`<a:videoFile r:link>`）這種現代 Word 的做法；若 docx 內以 OLE 物件（`<w:object>`／`word/embeddings/*.bin`）嵌入影片（少數舊版做法），不會被偵測到、也不會被保留。
- DOMParser 解析失敗時退回的 `docxXmlToTextLegacy()` 安全網完全不支援媒體擷取（原本就有的限制，未強化）。
- 已用 python-docx 產生的測試檔（含一張圖片＋前後段落＋一張表格）端對端驗證：上傳後 `#rawInput` 正確顯示 `[[IMG_1]]` 佔位標記於正確位置 → 陽春模式整理後渲染成真正 `<img>` → 下載 Word 後用 python-docx 讀回 `inline_shapes` 確認圖片正確內嵌（含等比例縮放的正確尺寸）、表格與文字段落皆正確。影片路徑因手邊沒有真的內嵌本機影片的 docx 樣本，僅經程式碼審視、未實測。

## 部署

已推公開 GitHub repo `M255525/text-organizer-studio`，用 `.github/workflows/deploy-pages.yml`（比照 `scamper-thinking-generator` 逐字複製）以 Actions workflow 部署 GitHub Pages（非 legacy branch-source，`gh api repos/M255525/text-organizer-studio/pages -f build_type=workflow` 開啟），已上線：<https://m255525.github.io/text-organizer-studio/>。

## 已部署上線工具的標準配件（2026-09-08 補齊）

比照 `scamper-thinking-generator`/`ai-prompt-generator` 既有做法，全部逐字複製或直接沿用同一套邏輯，只改動品牌相關的字串（storage key、page_id、標題等）：

- **頂部跑馬燈**：`#marqueeBar` IIFE 逐字複製，共用工作區既有的公告 Google Apps Script 端點與 Sheet（與其餘工具同一顆，改內容不必重新部署），`localStorage` key `textOrganizerMarquee`。
- **使用警語＋創作者資訊**：`footer` 內 `.warn-box`（5 點警語，內容依本工具實際功能調整措辭）＋ `.footer-meta`（Mark Tsai 聯絡信箱、訪客計數器、加入主畫面按鈕、操作手冊連結）。
- **`manual.html`**：操作手冊，內容依本工具實際功能撰寫（貼上/上傳、API 設定與額外需求解鎖、AI 與陽春兩種整理模式、可編輯輸出區、表格保留、隱私與警語）；創作者資料段落與 `scamper-thinking-generator/manual.html` 等姊妹專案為同一份，更新其中一邊時同步其餘各邊。
- **訪客計數器**：`visitor-badge.laobi.icu`，`page_id=m255525.text-organizer-studio`。
- **PWA 加入主畫面**：`manifest.json`／`service-worker.js`（network-first + 同源快取備援）／`icons/`（`icon-192.png`／`icon-512.png`／`icon-maskable-512.png`／`apple-touch-icon.png`，navy 底＋金色「整」字，用 PIL 現畫、`msjhbd.ttc` 字型，產生後即刪除生成腳本）＋ 獨立 IIFE 安裝按鈕邏輯（iOS/macOS Safari 無 `beforeinstallprompt` 時顯示對應操作指引，安裝腳本自帶 `notify()` 不依賴主程式的 `showToast`，避免跨 IIFE 作用域看不到的既知坑）。已用 `navigator.serviceWorker.getRegistrations()` 驗證 SW 確實註冊並 active。
- 已用瀏覽器自動化端對端驗證：跑馬燈正確抓到共用 Sheet 內容並顯示、警語與創作者資訊正常渲染、訪客計數器正常顯示、manual.html 各段落正常、加入主畫面按鈕可點擊無錯誤。

## 本次刻意未做

- 序號授權（`member-license-gate`）——使用者未要求套用
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
