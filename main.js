const ss = SpreadsheetApp.getActiveSpreadsheet();
const parametersSheet = ss.getSheetByName("參數區");
const prearrangedSheet = ss.getSheetByName("預先編排的科目及節次");
const unfilteredSheet = ss.getSheetByName("註冊組補考名單");
const candidateSheet = ss.getSheetByName("教學組排入考程的科目");
const openSheet = ss.getSheetByName("開課資料(查詢任課教師用)");
const filteredSheet = ss.getSheetByName("排入考程的補考名單");
const smallBagSheet = ss.getSheetByName("小袋封面套印用資料");
const bigBagSheet = ss.getSheetByName("大袋封面套印用資料");
const bulletinSheet = ss.getSheetByName("公告版補考場次");
const sessionRecordSheet = ss.getSheetByName("試場紀錄表(A表)");

// 建立工作列選單
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu("補考節次試場編排小幫手")
        .addItem("註冊組補考名單課程代碼補完", "unfilted_code_complete")
        .addItem("開課資料課程代碼補完", "open_code_complete")
        .addSeparator()
        .addItem("步驟 1. 產出公告用補考名單、試場記錄表", "all_in_one")
        .addItem("步驟 2. 產生合併列印小袋封面(要很久哦)", "merge_to_small_bag")
        .addItem(
            "產生合併大袋封面用資料(人工輸入監考教師)",
            "generate_big_bag_data"
        )
        .addItem("步驟 3. 產生合併列印大袋封面", "merge_to_big_bag")
        .addSeparator()
        .addItem("依「科目」排序補考名單", "sort_by_subject")
        .addItem("依「班級座號」排序補考名單", "sort_by_classname")
        .addItem("依「節次試場」排序補考名單", "sort_by_session_classroom")
        .addSeparator()
        .addItem("1-1. 清空", "initialize")
        .addItem("1-2. 開始篩選", "getFilteredData")
        .addItem("1-3. 安排共同科節次", "arrange_commons_session")
        .addItem("1-4. 安排專業科節次", "arrangeProfessionsSession")
        .addItem("1-5. 安排試場", "arrangeClassroom")
        .addItem("1-6. 計算大、小袋編號", "bag_numbering")
        .addItem("1-7. 計算試場人數", "calculate_classroom_population")
        .addItem("1-8. 產生「公告版補考場次」", "generate_bulletin")
        .addItem("1-9. 產生「試場記錄表」", "generate_sessionRecordSheet")
        .addItem("1-10. 產生「小袋封面套印用資料」", "generate_small_bag_data")
        .addItem("1-11. 產生「大袋封面套印用資料」", "generate_big_bag_data")
        .addToUi();
}

function initialize() {
    // 將工作表「排入考程的補考名單」初始化成只剩第一列的欄位標題
    // (1) 清除所有儲存格內容
    // (2) 刪除多餘的列到只剩5列
    // (3) 填入欄位標題

    // 清除所有值
    filteredSheet.clear();
    smallBagSheet.clear();
    bigBagSheet.clear();
    bulletinSheet.clear();
    sessionRecordSheet.clear();

    // 將課程代碼補完，包括：「註冊組匯出的補考名單」、「開課資料(查詢任課教師用)」
    unfilted_code_complete();
    open_code_complete();

    // 清空資料並設置標題列
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
    filteredSheet.clear();
    filteredSheet.appendRow(headers);

    // 移除已有篩選器，重新設置新的篩選器
    if (filteredSheet.getDataRange().getFilter()) {
        filteredSheet.getDataRange().getFilter().remove();
    }
    filteredSheet.getDataRange().createFilter();
}

// 一鍵產出公告用補考名單、試場記錄表
function all_in_one() {
    // Start counting execution time
    let runtime_count_start = new Date();

    getFilteredData(); // 篩選出列入考程的科目
    arrange_commons_session(); // 安排物理、國、英、數、資訊科技、史地的節次
    arrangeProfessionsSession(); // 安排專業科目的節次
    arrangeClassroom(); // 安排試場的班級科目
    sort_by_session_classroom();
    bag_numbering();
    set_session_time();
    calculate_classroom_population();
    generate_bulletin();
    generate_sessionRecordSheet();

    // Stop counting execution time
    newRuntime = runtime_count_stop(runtime_count_start);

    SpreadsheetApp.getUi().alert("已完成編排，共使用" + newRuntime + "秒");
}
