function get_filtered_data() {
    // 取得「註冊組補考名單」的欄位索引
    const [unfilteredSheetHeaders, ...unfiltered_data] = unfilteredSheet
        .getDataRange()
        .getValues();
    const std_number_column = unfilteredSheetHeaders.indexOf("學號");
    const class_column = unfilteredSheetHeaders.indexOf("班級");
    const seat_number_column = unfilteredSheetHeaders.indexOf("座號");
    const std_name_column = unfilteredSheetHeaders.indexOf("姓名");
    const subject_name_column = unfilteredSheetHeaders.indexOf("科目名稱");
    const code_column = unfilteredSheetHeaders.indexOf("科目代碼補完");

    // 取得「開課資料」的欄位索引
    const [openSheet_headers, ...open_data] = openSheet
        .getDataRange()
        .getValues();
    const open_class_column = openSheet_headers.indexOf("班級名稱");
    const open_subject_name_column = openSheet_headers.indexOf("科目名稱");
    const teacher_column = openSheet_headers.indexOf("任課教師");

    // 取得「教學組排入考程的科目」的欄位索引
    const [candidate_subject_headers, ...candidate_subjects_data] =
        candidateSheet.getDataRange().getValues();
    const make_up_column = candidate_subject_headers.indexOf("要補考");
    const filtered_code_column = candidate_subject_headers.indexOf("課程代碼");
    const isComputerGradedColumn = candidate_subject_headers.indexOf("電腦");
    const isManualGradedColumn = candidate_subject_headers.indexOf("人工");

    let candidate_subjects = {};
    candidate_subjects_data.forEach(function (row) {
        if (row[make_up_column] == true) {
            candidate_subjects[row[filtered_code_column]] = {
                isComputerGraded: row[isComputerGradedColumn],
                isManualGraded: row[isManualGradedColumn],
            };
        }
    });

    let open_teacher = {};
    open_data.forEach(function (row) {
        if (row[teacher_column].toString().length > 10) {
            open_teacher[
                row[open_class_column].toString() +
                    row[open_subject_name_column].toString()
            ] = row[teacher_column].toString().split(",")[0].slice(7);
        } else {
            open_teacher[
                row[open_class_column].toString() +
                    row[open_subject_name_column].toString()
            ] = row[teacher_column].toString().slice(6);
        }
    });

    // 清除「排入考程的補考名單」內容
    initialize();

    let nameList = [];
    unfiltered_data.forEach(function (row) {
        if (Object.keys(candidate_subjects).includes(row[code_column])) {
            // 科別	年級	班級代碼	班級	座號	學號	姓名	科目名稱	節次	試場	小袋序號	小袋人數	大袋序號	大袋人數	班級人數	時間	電腦	人工	任課老師
            let tmp = [
                getDepartmentName(row[class_column]), // 科別
                getGrade(row[class_column]), // 年級
                getClassCode(row[class_column]), // 班級代碼
                row[class_column], // 班級
                row[seat_number_column], // 座號
                row[std_number_column], // 學號
                row[std_name_column], // 姓名
                row[subject_name_column], // 科目名稱
                (row[8] = 0), // 節次預設為0
                (row[9] = 0), // 試場預設為0
                "", // 小袋序號
                "", // 小袋人數
                "", // 大袋序號
                "", // 大袋人數
                "", // 班級人數
                "", // 時間
                candidate_subjects[row[code_column]]["isComputerGraded"]
                    ? "☑"
                    : "☐", // 電腦
                candidate_subjects[row[code_column]]["isManualGraded"]
                    ? "☑"
                    : "☐", //人工
                open_teacher[
                    row[class_column].toString() +
                        row[subject_name_column].toString()
                ], // 任課老師
            ];

            nameList.push(tmp);
        }
    });

    filteredSheet
        .getRange(2, 1, nameList.length, nameList[0].length)
        .setNumberFormat("@STRING@") // 改成純文字格式，以免 0 開頭的學號被去掉前面的 0，造成位數錯誤
        .setValues(nameList);
}
