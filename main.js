const ss = SpreadsheetApp.getActiveSpreadsheet();
const parametersSheet = ss.getSheetByName("參數區");
const prearrangedSheet = ss.getSheetByName("預先編排的科目及節次");
const unfilteredSheet = ss.getSheetByName("註冊組達補考標準名單");
const candidateSheet = ss.getSheetByName("教學組排入考程的科目");
const openSheet = ss.getSheetByName("開課資料(查詢任課教師用)");
const filteredSheet = ss.getSheetByName("補考名單");
const smallBagSheet = ss.getSheetByName("小袋封面套印用資料");
const bigBagSheet = ss.getSheetByName("大袋封面套印用資料");
const bulletinSheet = ss.getSheetByName("公告版補考場次");
const sessionRecordSheet = ss.getSheetByName("試場紀錄表(A表)");
const configs = getConfigs();

// 建立工作列選單
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu("補考節次試場編排小幫手")
        .addItem("註冊組補考名單課程代碼補完", "completeUnfilteredSheetCode")
        .addItem("開課資料課程代碼補完", "completeOpenSheetCode")
        .addSeparator()
        .addItem(
            "步驟 1. 產出公告用補考名單、試場記錄表",
            "generateCompleteExamArrangement"
        )
        .addItem("步驟 2. 產生合併列印小袋封面(要很久哦)", "merge_to_small_bag")
        .addItem(
            "產生合併大袋封面用資料(人工輸入監考教師)",
            "generate_big_bag_data"
        )
        .addItem("步驟 3. 產生合併列印大袋封面", "merge_to_big_bag")
        .addSeparator()
        .addItem("依「科目」排序補考名單", "sortBySubject")
        .addItem("依「班級座號」排序補考名單", "sortByClassName")
        .addItem("依「節次試場」排序補考名單", "sortBySessionThenClassroom")
        .addSeparator()
        .addItem("1-1. 清空", "initialize")
        .addItem("1-2. 開始篩選", "getFilteredData")
        .addItem("1-3. 安排共同科目節次", "arrangeCommonSubjectSessions")
        .addItem("1-4. 安排專業科目節次", "arrangeProfessionsSession")
        .addItem("1-5. 安排試場", "arrangeClassroom")
        .addItem("1-6. 計算大、小袋編號", "bag_numbering")
        .addItem("1-7. 計算試場人數", "calculate_classroom_population")
        .addItem("1-8. 產生「公告版補考場次」", "generate_bulletin")
        .addItem("1-9. 產生「試場記錄表」", "generate_sessionRecordSheet")
        .addItem("1-10. 產生「小袋封面套印用資料」", "generate_small_bag_data")
        .addItem("1-11. 產生「大袋封面套印用資料」", "generate_big_bag_data")
        .addToUi();
}

/**
 * 刪除工作表多餘的列，只保留指定數量的列
 * @param {Sheet} sheet - 工作表物件
 * @param {number} keepRows - 要保留的列數
 * @returns {Sheet} 處理後的工作表
 */
const deleteExcessRows = (sheet, keepRows) => {
    const maxRows = sheet.getMaxRows();
    if (maxRows > keepRows) {
        sheet.deleteRows(keepRows + 1, maxRows - keepRows);
    }
    return sheet;
};

/**
 * 設置工作表範圍為純文字格式
 * @param {Sheet} sheet - 工作表物件
 * @param {number} rows - 列數
 * @param {number} columns - 欄數
 * @returns {Sheet} 處理後的工作表
 */
const setTextFormat = (sheet, rows, columns) => {
    sheet.getRange(1, 1, rows, columns).setNumberFormat("@STRING@");
    return sheet;
};

/**
 * 初始化單一工作表
 * @param {Sheet} sheet - 工作表物件
 * @param {number} columns - 欄數
 * @param {number} keepRows - 要保留的列數
 * @returns {Sheet} 處理後的工作表
 */
const initializeSheet = (sheet, columns, keepRows = 5) => {
    return [sheet]
        .map((s) => {
            s.clear();
            return s;
        })
        .map((s) => deleteExcessRows(s, keepRows))
        .map((s) => setTextFormat(s, keepRows, columns))[0];
};

/**
 * 取得工作表設定對應表
 * @returns {Array} 工作表設定陣列 [sheet, columns]
 */
const getSheetConfigurations = () => [
    [filteredSheet, 19],
    [bulletinSheet, 6],
    [sessionRecordSheet, 12],
    [smallBagSheet, 12],
    [bigBagSheet, 10],
];

/**
 * 批次初始化所有工作表
 * @param {Array} sheetConfigs - 工作表設定陣列
 * @returns {Array} 初始化後的工作表陣列
 */
const initializeAllSheets = (sheetConfigs) =>
    sheetConfigs.map(([sheet, columns]) => initializeSheet(sheet, columns));

/**
 * 設置補考名單標題列
 * @returns {void}
 */
const setupFilteredSheetHeaders = () => {
    const headers = [
        "科別",
        "年級",
        "班級代碼",
        "班級",
        "座號",
        "學號",
        "姓名",
        "科目名稱",
        "節次",
        "試場",
        "小袋序號",
        "小袋人數",
        "大袋序號",
        "大袋人數",
        "班級人數",
        "時間",
        "電腦",
        "人工",
        "任課老師",
    ];
    filteredSheet.appendRow(headers);
};

/**
 * 設置篩選器
 * @returns {void}
 */
const setupFilter = () => {
    if (filteredSheet.getDataRange().getFilter()) {
        filteredSheet.getDataRange().getFilter().remove();
    }
    filteredSheet.getDataRange().createFilter();
};

/**
 * 執行課程代碼補完
 * @returns {void}
 */
const completeCourseCodes = () => {
    completeUnfilteredSheetCode();
    completeOpenSheetCode();
};

/**
 * 將工作表「排入考程的補考名單」初始化成只剩第一列的欄位標題
 * (1) 清除所有儲存格內容
 * (2) 刪除多餘的列到只剩5列
 * (3) 填入欄位標題
 * @returns {void}
 */
function initialize() {
    // 批次初始化所有工作表
    const sheetConfigs = getSheetConfigurations();
    initializeAllSheets(sheetConfigs);

    Logger.log(
        "(initialize) 清除「排入考程的補考名單」、「公告版」、「試場紀錄表」、「大小袋封面」工作表的內容，刪減列數至5列，並設置為純文字格式"
    );

    // 補完課程代碼
    completeCourseCodes();
    Logger.log(
        "(initialize) 補完「註冊組補考名單」、「開課資料(查詢任課教師用)」工作表的課程代碼。"
    );

    // 設置補考名單標題和篩選器
    setupFilteredSheetHeaders();
    setupFilter();
}

/**
 * 一鍵產出公告用補考名單、試場記錄表
 * @returns {void}
 */
function generateCompleteExamArrangement() {
    const startTime = new Date();

    getFilteredData(); // 篩選出列入考程的科目
    arrangeCommonSubjectSessions(); // 安排物理、國、英、數、資訊科技、史地的節次
    arrangeProfessionsSession(); // 安排專業科目的節次
    arrangeClassroom(); // 安排試場的班級科目
    sortBySessionThenClassroom();
    bag_numbering();
    set_session_time();
    calculate_classroom_population();
    generate_bulletin();
    generate_sessionRecordSheet();

    newRuntime = calculateElapsedTimeInSeconds(startTime);

    SpreadsheetApp.getUi().alert(
        "(generateCompleteExamArrangement) 已完成補考場次編排，共使用" +
            newRuntime +
            "秒"
    );
}
