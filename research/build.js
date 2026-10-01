const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  HeadingLevel, AlignmentType, ImageRun, LevelFormat, BorderStyle, Footer, PageNumber, TableOfContents, PageBreak } = require('docx');

const FONT = 'Microsoft JhengHei';
const W = 9638; // A4 content width (DXA) with 2cm margins
const NAVY = '1F3864';

const p = (text, opts = {}) => new Paragraph({ spacing: { after: 120, line: 320 }, ...opts,
  children: Array.isArray(text) ? text : [new TextRun({ text, ...(opts.run || {}) })] });
const b = (t) => new TextRun({ text: t, bold: true });
const r = (t) => new TextRun({ text: t });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const bl = (runs, level = 0) => new Paragraph({ numbering: { reference: 'bul', level }, spacing: { after: 80, line: 300 },
  children: typeof runs === 'string' ? [r(runs)] : runs });
const kv = (k, v) => bl([b(k), r(v)]);

function table(headers, rows, widths) {
  const total = widths.reduce((a, c) => a + c, 0);
  const scale = W / total; const ws = widths.map(x => Math.round(x * scale));
  ws[ws.length - 1] += W - ws.reduce((a, c) => a + c, 0);
  const cell = (t, i, head, shade) => new TableCell({ width: { size: ws[i], type: WidthType.DXA },
    shading: head ? { type: ShadingType.CLEAR, fill: NAVY, color: 'auto' } : (shade ? { type: ShadingType.CLEAR, fill: 'F2F2F2', color: 'auto' } : undefined),
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    children: String(t).split('|').map(line => new Paragraph({ spacing: { after: 0, line: 270 },
      children: [new TextRun({ text: line, bold: head, color: head ? 'FFFFFF' : undefined, size: 18 })] })) });
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ws,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, i, true)) }),
      ...rows.map((row, ri) => new TableRow({ children: row.map((c, i) => cell(c, i, false, ri % 2 === 1)) }))] });
}
const gap = () => new Paragraph({ spacing: { after: 60 }, children: [] });
const note = (t) => p([new TextRun({ text: t, italics: true, size: 18, color: '666666' })]);

const C = [];
// ---------- Cover ----------
C.push(new Paragraph({ spacing: { before: 2400, after: 200 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'AIoT 產業鏈上下游與技術趨勢研究', bold: true, size: 48, color: NAVY })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 },
  children: [new TextRun({ text: 'Edge AI 驅動的第二成長曲線：價值鏈、利潤池、台廠定位與投資觀察架構', size: 26, color: '444444' })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '研究日期：2026 年 10 月 1 日', size: 22 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '幣別：除特別標註外，市場規模以 US$ 表示', size: 20, color: '666666' })] }));
C.push(new Paragraph({ children: [new PageBreak()] }));
C.push(new TableOfContents('目錄', { hyperlink: true, headingStyleRange: '1-2' }));
C.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 1. Key conclusions ----------
C.push(h1('一、核心結論'));
C.push(table(['#', '結論', '投資含義'], [
  ['1', 'AIoT 的本質是「推論從雲端下放到終端」，2025–26 年起 MCU/SoC 內建 NPU 已成標配，競爭軸從「連網」轉向「每瓦算力 (TOPS/W)」與軟體生態。', '價值往上游 SoC/MCU + NPU IP 集中；純連網晶片面臨 ASP 下滑。'],
  ['2', '利潤池呈微笑曲線：上游 IC/IP（GM 45–65%）與下游平台/訂閱服務（GM 60%+）最肥；中游模組/ODM（GM 10–25%）最薄。', '中游只挑有垂直領域 know-how 的工業電腦 / 安防品牌（GM 35%+）。'],
  ['3', '市場規模研調差異極大（2026 年 US$17–99B，CAGR 14–38%），原因是 AIoT 定義不一。真正可驗證的是硬體量：連網裝置 2025 年底約 211 億台，年增 ~14%。', '別用市場報告 CAGR 當估值依據；用個股營收、Design-win 數與 NPU 產品營收占比追蹤。'],
  ['4', '三條確定性最高的主線：(a) Edge AI MCU/SoC 換代；(b) Wi-Fi 7 / Matter / RedCap 連網升級；(c) 非紅色供應鏈（NDAA、美國 Cyber Trust Mark、EU CRA 2027 生效）。', '台廠受惠：聯發科、瑞昱、新唐、晶心科、力旺、華邦電、研華、晶睿、啟碁。'],
  ['5', '最大風險：中國廠（樂鑫、全志、瑞芯微、晶晨、移遠）價格戰 + 消費性需求疲弱；工業端庫存循環。', 'AIoT 不是 AI Server，成長是「慢而廣」，估值不應套用 AI Server 倍數。'],
], [4, 52, 44]));
C.push(gap());
C.push(note('判斷前提：2026 下半年全球消費電子需求溫和、工業 MCU 庫存已在 2025 年完成去化；若此假設不成立，結論 1、4 的時程需遞延 2–3 季。'));

// ---------- 2. Definition & size ----------
C.push(h1('二、產業定義與市場規模'));
C.push(h2('2.1 定義'));
C.push(p('AIoT（Artificial Intelligence of Things）= IoT 的感測/連網能力 + AI 的推論/決策能力。關鍵分野在於「AI 在哪裡跑」：'));
C.push(table(['架構', '運算位置', '典型算力', '代表應用', '主要晶片'], [
  ['Cloud AI + IoT', '雲端資料中心', 'PFLOPS 級', '大數據分析、模型訓練', 'GPU / ASIC'],
  ['Edge AI', '閘道器 / 邊緣伺服器 / IPC', '10–200 TOPS', 'AOI、影像分析、AMR', 'Jetson、Genio、x86+NPU'],
  ['Endpoint / TinyML', '終端裝置本身', '0.1–2 TOPS', '喚醒詞、手勢、異常偵測', 'MCU + NPU（Ethos-U、Neural-ART）'],
], [18, 22, 15, 25, 20]));
C.push(h2('2.2 市場規模（研調比較）'));
C.push(table(['來源', '2026 規模', '目標年 / 規模', 'CAGR', '口徑說明'], [
  ['Research and Markets', 'US$17.5B', '—', '32.6%', '窄口徑：AIoT 解決方案'],
  ['Fortune Business Insights', 'US$82.7B', '2034 / US$781B', '32.4%', '含硬體、軟體、服務'],
  ['Mordor Intelligence (AI in IoT)', 'US$74–99B', '2031–32', '14–22%', '中性口徑'],
  ['其他研調', 'US$75.5B', '2033 / US$724B', '38.1%', '偏樂觀'],
  ['IoT Analytics（裝置數）', '2025 年底 211 億台', '2030 年 ~400 億台', '~14%（台數）', '最可驗證的硬體量指標'],
], [26, 16, 20, 12, 26]));
C.push(gap());
C.push(p([b('解讀：'), r('同年度規模差距達 5 倍，代表研調數字僅適合判斷方向，不適合作為營收預估基礎。較可信的推估方式是「裝置數成長（~14%）× AI 滲透率提升 × 單機 ASP 上升」，三者相乘得出 AIoT 晶片端合理成長約 20–25%/年，此為本報告的基準假設。')]));

// ---------- 3. Chain ----------
C.push(h1('三、產業鏈上下游拆解'));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ type: 'jpg', data: fs.readFileSync('chain.jpg'),
  transformation: { width: 640, height: 304 } })] }));
C.push(h2('3.1 上游：晶片、元件與矽智財'));
C.push(table(['次領域', '技術關鍵', '國際龍頭', '台廠代表（代號）', '典型 GM'], [
  ['AIoT SoC', 'CPU+NPU+ISP 整合、6nm–12nm', 'Qualcomm、NXP、Ambarella', '聯發科 Genio (2454)、瑞昱 (2379)、凌陽 (2401)', '45–55%'],
  ['MCU + NPU', 'Cortex-M55/M85 + Ethos-U、TinyML', 'ST、Renesas、NXP、Infineon', '新唐 M55M1 (4919)、盛群 (6202)', '40–50%'],
  ['連網晶片', 'Wi-Fi 7、BLE 6、Thread/Matter', 'Broadcom、Qualcomm、樂鑫', '瑞昱、聯發科 Filogic', '40–50%'],
  ['影像 / 感測', 'CMOS、ToF、光學感測、毫米波雷達', 'Sony、豪威、TI', '原相 (3227)、義隆 (2458)', '40–55%'],
  ['記憶體', 'NOR/SLC NAND、客製化 CUBE', 'Winbond、Macronix、GigaDevice', '華邦電 (2344)、旺宏 (2337)、鈺創 (5351)', '20–45%（循環）'],
  ['矽智財 IP', 'RISC-V CPU、NPU IP、安全/OTP', 'Arm、SiFive、Synopsys', '晶心科 (6533)、力旺 (3529)', '90%+'],
  ['晶圓代工', '22/28/40nm eNVM、BCD', 'TSMC、GF', '聯電 (2303)、世界 (5347)', '25–45%'],
], [13, 24, 20, 30, 13]));
C.push(gap());
C.push(p([b('重點：'), r('上游最具定價權的是「整合型 SoC + 軟體工具鏈」。客戶選擇 Edge AI 晶片時，NPU 峰值 TOPS 只是門票，真正決定 Design-win 的是模型轉換工具、SDK 成熟度與長期供貨保證（工業客戶要求 10–15 年）。')]));
C.push(h2('3.2 中游：模組、設備與平台'));
C.push(table(['次領域', '角色', '國際/中國對手', '台廠代表（代號）', '典型 GM'], [
  ['連網 / AI 模組', '晶片封裝成可認證模組', '移遠、廣和通、u-blox', '環旭 (6231)、正文 (4906)', '10–20%'],
  ['網通 CPE / 閘道器', 'Wi-Fi 7 Router、FWA、Mesh', 'TP-Link、華為', '啟碁 (6285)、中磊 (5388)、智易 (3596)', '10–15%'],
  ['工業電腦 / 邊緣運算', 'Edge AI Box、IPC、強固型平板', 'Siemens、Kontron', '研華 (2395)、凌華 (6166)、樺漢 (6414)、振樺電 (8114)', '25–40%'],
  ['工業儲存', '寬溫 SSD、DRAM 模組、AI 加速卡', 'Swissbit', '宜鼎 (5289)、宇瞻 (8271)', '25–35%'],
  ['安防影像', 'AI 攝影機、NVR、VMS', '海康、大華', '晶睿 (3454)、奇偶 (3356)', '45–50%'],
  ['ODM / EMS', '終端組裝', '立訊、比亞迪電子', '鴻海 (2317)、和碩 (4938)', '5–10%'],
  ['平台 / 軟體', '裝置管理、OTA、資安、數位孿生', 'AWS IoT、Azure IoT、PTC', '研華 WISE-IoT（少數）', '60%+'],
], [15, 24, 18, 31, 12]));
C.push(gap());
C.push(p([b('重點：'), r('中游普遍被上下游擠壓，唯二例外：(1) 工業電腦 —— 少量多樣、認證門檻與長生命週期造成黏著度；(2) 非中國安防品牌 —— 受惠美國 NDAA §889 與 FCC 對海康/大華的禁令，具地緣溢價。')]));
C.push(h2('3.3 下游：應用場景'));
C.push(table(['應用', '2026–28 成長動能', 'AI 功能', '景氣敏感度', '對台廠受惠度'], [
  ['工業 4.0 / 智慧工廠', '高', 'AOI 瑕疵檢測、預測性維護、AMR', '中（隨 capex）', '高（IPC、MCU）'],
  ['智慧家庭', '中', '語音、人體感測、Matter 互通', '高（消費性）', '中（Wi-Fi/BLE IC）'],
  ['智慧安防 / 城市', '中高', '人車辨識、VLM 影像搜尋', '低（政府標案）', '高（晶睿）'],
  ['AI 穿戴 / AI 眼鏡', '高（低基期）', '端側 SLM、即時翻譯、視覺辨識', '高', '中（SoC、光學、組裝）'],
  ['Physical AI / 機器人', '極高（早期）', 'VLA 模型、感測融合、即時控制', '中', '中高（關節模組、IPC、感測）'],
  ['車聯網 / 智慧座艙', '中', 'DMS、語音助理、V2X', '中', '中（聯發科 Dimensity Auto）'],
  ['能源 / 智慧電網', '中高', '負載預測、BESS 能源管理', '低（政策驅動）', '中（MCU、BMS）'],
  ['智慧醫療 / 零售', '中', '遠距照護、無人商店、ESL', '低', '低中'],
], [18, 16, 30, 16, 20]));

// ---------- 4. Profit pool ----------
C.push(h1('四、價值分布與利潤池'));
C.push(table(['環節', '價值占比（估）', '毛利率', '競爭強度', '定價權', '結論'], [
  ['晶片 / IP', '20–25%', '45–65%（IP 90%+）', '中（寡占）', '高', '核心利潤池'],
  ['模組 / ODM / EMS', '25–30%', '5–20%', '極高（中國價格戰）', '低', '避開或只挑利基'],
  ['工業電腦 / 垂直設備', '10–15%', '25–45%', '中', '中高', '選擇性布局'],
  ['連網服務 / 電信', '10%', '30–40%', '中', '中', '台股參與度低'],
  ['平台軟體 / SI / 訂閱', '25–30%', '60%+', '中（雲端巨頭）', '高', '美股為主'],
], [20, 14, 18, 18, 10, 20]));
C.push(gap());
C.push(note('價值占比為作者依產業結構估算之區間，非研調精確數據，僅作相對比較用途。'));

// ---------- 5. Tech trends ----------
C.push(h1('五、八大技術趨勢'));
const trends = [
  ['T1. MCU/SoC 全面 NPU 化', [
    ['現況：', 'ST STM32N6（Neural-ART，600 GOPS）、Renesas RA8P1（Ethos-U55，256 GOPS）、NXP i.MX RT700（eIQ Neutron）、新唐 M55M1（Ethos-U55，~128 GOPS）已量產。'],
    ['方向：', 'MCU 端 0.1–1 TOPS、MPU/SoC 端 5–50 TOPS；AI PC/高階邊緣的門檻已推升至 50 TOPS。'],
    ['影響：', '單顆 MCU ASP 由 US$1–3 提升至 US$4–10，量不必大增即可推升營收；Arm Ethos 生態與 RISC-V 自研 NPU 兩路線並行。']]],
  ['T2. 生成式 AI 下放終端（端側 SLM / VLM）', [
    ['現況：', '1–4B 參數小型語言模型經 INT4 量化可在 8–16GB LPDDR5X 的邊緣 SoC 上運行；安防攝影機導入 VLM 做自然語言影像搜尋。'],
    ['影響：', '終端記憶體容量與頻寬成新瓶頸 → 帶動 LPDDR、客製化記憶體（如華邦 CUBE）需求；Edge AI Box 從 CNN 推論升級為 Transformer 推論。']]],
  ['T3. 連網標準升級', [
    ['Wi-Fi 7：', '2025–27 年滲透率快速提升，MLO 降低延遲，路由器/閘道器換機潮。'],
    ['Matter / Thread：', '統一智慧家庭互通，降低品牌鎖定，有利晶片廠出貨、不利封閉生態品牌。'],
    ['5G RedCap / LTE Cat-1bis：', '取代 4G Cat-1/Cat-4，用於穿戴、工業感測、POS。'],
    ['NTN（衛星直連 IoT）：', '3GPP Rel-17/18，資產追蹤、海事、農業，2026–28 開始商轉。']]],
  ['T4. RISC-V 滲透', [
    ['現況：', 'IoT/MCU 是 RISC-V 最先規模化的市場，中國廠為規避 Arm 授權與出口管制大量採用。'],
    ['影響：', '晶心科（6533）受惠授權+權利金；但 RISC-V 也降低 MCU 進入門檻，加劇價格競爭。']]],
  ['T5. 特殊製程與先進封裝下沉', [
    ['方向：', 'AIoT SoC 主力在 6–12nm；MCU 在 22/28/40nm eNVM（eFlash→MRAM/RRAM）。小晶片（Chiplet）與 SiP 開始用於 AI 眼鏡與穿戴，以縮小體積。'],
    ['影響：', '聯電、世界先進的 22/28nm 特殊製程產能利用率受益；SiP 有利環旭。']]],
  ['T6. 超低功耗與能源採集', [
    ['方向：', 'Always-on AI 需 µW 級待機；能源採集（光、熱、振動）+ 無電池感測器；BLE 6 Channel Sounding 提供精準定位。'],
    ['影響：', '類比/電源 IC 與 PMIC 重要性提升。']]],
  ['T7. 資安法規強制化', [
    ['EU Cyber Resilience Act：', '2027 年 12 月全面適用，聯網產品須具備安全更新與漏洞通報義務。'],
    ['美國 Cyber Trust Mark、NDAA §889：', '強化非中國供應鏈需求。'],
    ['影響：', '安全 IP（力旺 PUF/OTP）、安全元件、具資安認證的晶片與設備具溢價；中國模組/攝影機在歐美市場受限。']]],
  ['T8. Physical AI 與數位孿生', [
    ['方向：', 'AIoT 從「感知」走向「行動」：AMR、人形機器人、自主產線。NVIDIA Jetson Thor / Omniverse 建立模擬到部署的標準流程。'],
    ['影響：', '對 IPC、感測、伺服/關節模組、高速連接器的需求擴大，是 2027 年後最大的潛在增量，但目前營收貢獻仍小。']]],
];
for (const [title, items] of trends) {
  C.push(h2(title));
  for (const [k, v] of items) C.push(kv(k, v));
}
C.push(h2('技術成熟度總表'));
C.push(table(['趨勢', '成熟度', '主要營收貢獻期', '確定性', '最受惠環節'], [
  ['T1 NPU 化 MCU/SoC', '量產擴散期', '2025–2028', '高', '上游 IC'],
  ['T2 端側生成式 AI', '導入期', '2026–2029', '中高', 'SoC、記憶體'],
  ['T3 Wi-Fi 7 / Matter / RedCap', '量產擴散期', '2025–2027', '高', '連網 IC、網通'],
  ['T4 RISC-V', '成長期', '2025–2030', '中高', 'IP'],
  ['T5 特殊製程 / SiP', '成熟期', '持續', '中', '晶圓代工、SiP'],
  ['T6 超低功耗', '成長期', '2026–2029', '中', '類比 IC'],
  ['T7 資安法規', '法規驅動', '2026–2028', '高', '安全 IP、非紅供應鏈'],
  ['T8 Physical AI', '早期', '2027–2032', '中（時程不確定）', 'IPC、感測、零組件'],
], [26, 16, 18, 18, 22]));

// ---------- 6. Competition ----------
C.push(h1('六、競爭格局'));
C.push(table(['陣營', '代表廠商', '優勢', '劣勢', '對台廠影響'], [
  ['歐美日 IDM', 'ST、NXP、Renesas、Infineon、TI、Microchip', '工業/車用認證、長期供貨、自有晶圓廠', '價格高、庫存循環大', '台廠以性價比切中階市場'],
  ['美系平台', 'Qualcomm、NVIDIA、Ambarella', '高階算力、軟體生態', '價格高，低階覆蓋弱', '聯發科 Genio 正面競爭'],
  ['中國廠', '樂鑫、全志、瑞芯微、晶晨、移遠、海康', '成本、本土市場、速度', '歐美地緣限制、資安疑慮', '中低階價格戰主要壓力來源'],
  ['台灣廠', '聯發科、瑞昱、新唐、研華、晶睿', '非紅供應鏈、成本與品質平衡', '品牌與軟體平台弱', 'China+1 主要受惠者'],
], [14, 26, 22, 20, 18]));
C.push(gap());
C.push(p([b('格局判斷：'), r('AIoT 市場呈「高階美系、中低階中國、台灣卡位中間 + 非紅供應鏈」的三層結構。台廠的核心護城河不是技術領先，而是「地緣可信度 × 成本」。這也意味若中美關係顯著緩和，台廠溢價會被壓縮。')]));

// ---------- 7. TW map ----------
C.push(h1('七、台股 AIoT 供應鏈地圖'));
C.push(table(['公司（代號）', '環節', 'AIoT 角色', '觀察重點', '屬性'], [
  ['聯發科 (2454)', '上游 SoC', 'Genio 系列、Filogic Wi-Fi 7、Dimensity Auto', 'AIoT 占營收比、Edge AI design-win', '龍頭 / 穩健'],
  ['瑞昱 (2379)', '上游 SoC/連網', 'AmebaPro2（含 NPU）、Wi-Fi/BLE、網通晶片', 'Wi-Fi 7 滲透、IPCam SoC 出貨', '穩健'],
  ['新唐 (4919)', '上游 MCU', 'M55M1（Ethos-U55）、安全 MCU', '日本子公司車用回溫、NPU MCU 量產', '景氣循環'],
  ['盛群 (6202)', '上游 MCU', '家電/健康量測 MCU', '中國家電需求、價格競爭', '景氣循環'],
  ['原相 (3227)', '上游感測', '光學/影像感測、AI 感測方案', '新應用（穿戴、機器人）營收占比', '轉型'],
  ['晶心科 (6533)', '上游 IP', 'RISC-V CPU IP', '權利金成長、AI/向量核授權', '高 beta'],
  ['力旺 (3529)', '上游 IP', 'OTP/PUF 安全 IP', 'CRA 法規帶動安全 IP 授權', '高估值'],
  ['華邦電 (2344)', '上游記憶體', 'NOR、SLC NAND、CUBE 客製化記憶體', 'CUBE 量產時程、利基型 DRAM 報價', '景氣循環'],
  ['研華 (2395)', '中游 IPC', 'Edge AI 平台、WISE-IoT', '接單 B/B ratio、Edge AI 營收占比', '龍頭 / 高品質'],
  ['凌華 (6166)', '中游 IPC', 'Edge AI 運算、機器人控制器', '虧轉盈進度', '轉機'],
  ['樺漢 (6414)', '中游 IPC', '工業電腦、能源 / POS', '併購整合綜效', '價值型'],
  ['宜鼎 (5289)', '中游儲存', '工業 SSD、Edge AI 方案', 'AI 解決方案營收占比', '成長'],
  ['晶睿 (3454)', '中游安防', 'AI 網路攝影機、VMS', '歐美非紅需求、軟體訂閱', '地緣受惠'],
  ['啟碁 (6285)', '中游網通', 'Wi-Fi 7、衛星、車用連網', 'LEO 衛星與 Wi-Fi 7 出貨', '題材多元'],
  ['中磊 (5388) / 智易 (3596)', '中游網通', 'CPE、FWA、Wi-Fi 7 閘道器', '電信商資本支出、毛利率', '代工屬性'],
  ['環旭 (6231)', '中游 SiP/EMS', '穿戴 SiP、AI 眼鏡模組', 'AI 穿戴新品量產', '量大利薄'],
], [18, 12, 30, 26, 14]));
C.push(gap());
C.push(note('本表僅為產業定位整理，未納入即時股價與財務預估；個股估值請另以最新財報與股價（Kevin 股價模型 / 台股雙法）計算。'));

// ---------- 8. Risks ----------
C.push(h1('八、風險評估'));
C.push(table(['風險', '發生機率', '影響程度', '說明', '觀察指標'], [
  ['中國價格戰', '高', '高', '樂鑫、全志等在中低階 Wi-Fi/MCU 持續殺價', '台廠 MCU/連網 IC 毛利率'],
  ['消費性需求疲弱', '中', '中高', '智慧家庭、穿戴受消費景氣與關稅影響', '美國零售銷售、家電出貨'],
  ['工業庫存循環', '中', '高', 'IDM 庫存反轉時拉低全產業 ASP', 'ST/NXP/Renesas 庫存天數、研華 B/B'],
  ['Edge AI 應用落地慢於預期', '中', '中', '缺乏殺手級應用，NPU 成為「規格」而非「營收」', 'NPU 產品營收占比'],
  ['地緣政治緩和', '低', '中', '非紅供應鏈溢價縮小', '美中貿易政策、NDAA 執行'],
  ['關稅 / 匯率', '中', '中', '美國關稅政策與 NT$ 升值侵蝕出口毛利', 'NT$/US$、Section 232/301'],
  ['記憶體價格上漲', '中高', '中', 'AI Server 排擠產能，墊高 AIoT 裝置 BOM', 'DRAM/NAND 合約價'],
], [18, 10, 10, 34, 28]));

// ---------- 9. Framework ----------
C.push(h1('九、投資觀察架構'));
C.push(h2('9.1 選股濾網'));
C.push(bl([b('定價權：'), r('毛利率 > 40%（IC）或 > 30%（設備），且過去 8 季標準差小。')]));
C.push(bl([b('AI 曝險可驗證：'), r('公司揭露 Edge AI / NPU 產品營收占比或 design-win 數，而非僅新聞稿題材。')]));
C.push(bl([b('非紅供應鏈溢價：'), r('歐美營收占比 > 40%。')]));
C.push(bl([b('景氣位置：'), r('庫存天數回落至正常區間、月營收 YoY 轉正並連續 3 個月以上。')]));
C.push(h2('9.2 估值原則'));
C.push(table(['類型', '適用估值法', '合理區間（參考）', '說明'], [
  ['IC 設計（AIoT SoC/MCU）', 'Forward P/E、PEG', 'P/E 15–25x', '景氣循環股須用循環中段 EPS，避免高峰 EPS 低 P/E 陷阱'],
  ['矽智財', 'EV/Sales、P/E', 'P/E 30–50x', '高毛利、權利金遞延，對成長率極敏感'],
  ['工業電腦', 'Forward P/E', 'P/E 18–28x', '龍頭品質溢價；轉機股看 P/B'],
  ['網通 / EMS', 'P/E、殖利率', 'P/E 10–15x', '代工屬性，不宜給 AI 溢價'],
  ['記憶體', 'P/B', 'P/B 1.0–2.0x', '循環谷底買 P/B'],
], [24, 20, 18, 38]));
C.push(gap());
C.push(note('估值區間為台股同類公司歷史常見範圍之定性參考，並非目標價。'));
C.push(h2('9.3 追蹤時程（2026Q4–2027）'));
C.push(bl('每月 10 日：台廠月營收（新唐、瑞昱、研華、晶睿 YoY）'));
C.push(bl('2027 年 1 月 CES：AI 眼鏡、Matter、Edge AI PC/Box 新品'));
C.push(bl('2027 年 2–3 月 Embedded World（紐倫堡）：MCU/NPU 新品與工業 Edge AI 方案'));
C.push(bl('2027 年 3 月 Computex 前夕 / NVIDIA GTC：Jetson 與 Physical AI 平台更新'));
C.push(bl('ST、NXP、Renesas、Infineon 季報：工業 MCU 庫存與訂單能見度'));
C.push(bl('EU CRA 2027 年 12 月全面適用前之認證進度'));

// ---------- 10. Sources & disclaimer ----------
C.push(h1('十、資料來源'));
for (const s of [
  'Research and Markets, Artificial Intelligence of Things (AIoT) Market Report 2026',
  'Fortune Business Insights, AIoT Market Size, Share, 2034',
  'Mordor Intelligence, AI in IoT Market',
  'KaaIoT, IoT and AI in 2026 — market scale, architecture, and industry impact',
  'IoT Insider / Wireless Logic, Global IoT connections set to reach 21.9 billion by 2026',
  'Axis Intelligence, IoT Statistics 2026（引用 IoT Analytics：2025 年底 211 億台裝置）',
  'MakerPro, 2026 Edge AI MCU 技術趨勢與廠商方案現況比較',
  'OmniXRI, 2026 台灣 AI 晶片發展現況及未來挑戰',
  '電子工程專輯, 新唐科技 2026 新品發佈會',
]) C.push(bl(s));
C.push(h1('免責聲明'));
C.push(p('本報告為產業研究整理，所列公司僅為產業定位說明，非投資建議。市場規模數據引用公開研調，不同機構口徑差異大；毛利率與價值占比為產業典型區間之估算，非特定公司財務數據。投資人應自行查證最新財報與市場資訊並獨立判斷。'));

const doc = new Document({
  styles: {
    default: { document: { run: { font: { ascii: FONT, eastAsia: FONT, hAnsi: FONT }, size: 21 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, color: NAVY }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0,
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, color: '2E75B6' }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: 'bul', levels: [
    { level: 0, format: LevelFormat.BULLET, text: '▪', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } } } },
    { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 800, hanging: 260 } } } }] }] },
  features: { updateFields: true },
  sections: [{ properties: { page: { margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'AIoT 產業鏈研究｜非投資建議｜第 ', size: 16, color: '888888' }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '888888' }), new TextRun({ text: ' 頁', size: 16, color: '888888' })] })] }) },
    children: C }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync('AIoT產業鏈與技術趨勢研究_2026Q4.docx', buf); console.log('done'); });
