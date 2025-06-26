function getFilteredData() {
    // 取得「註冊組補考名單」的欄位索引
    const [unfilteredSheetHeaders, ...unfilteredData] = unfilteredSheet
        .getDataRange()
        .getValues();
    const studentIdColumn = unfilteredSheetHeaders.indexOf("學號");
    const classNameColumn = unfilteredSheetHeaders.indexOf("班級");
    const seatNumberColumn = unfilteredSheetHeaders.indexOf("座號");
    const studentNameColumn = unfilteredSheetHeaders.indexOf("姓名");
    const subjectNameColumn = unfilteredSheetHeaders.indexOf("科目名稱");
    const subjectCodeColumn = unfilteredSheetHeaders.indexOf("科目代碼補完");

    // 取得「開課資料」的欄位索引
    const [openSheetHeaders, ...openData] = openSheet
        .getDataRange()
        .getValues();
    const openClassNameColumn = openSheetHeaders.indexOf("班級名稱");
    const openSubjectNameColumn = openSheetHeaders.indexOf("科目名稱");
    const teacherNameColumn = openSheetHeaders.indexOf("任課教師");

    // 取得「教學組排入考程的科目」的欄位索引
    const [candidateSheetHeaders, ...candidateSheetData] = candidateSheet
        .getDataRange()
        .getValues();
    const makeUpColumn = candidateSheetHeaders.indexOf("要補考");
    const filteredSubjectCodeColumn = candidateSheetHeaders.indexOf("課程代碼");
    const isComputerGradedColumn = candidateSheetHeaders.indexOf("電腦");
    const isManualGradedColumn = candidateSheetHeaders.indexOf("人工");

    let candidateSubjects = {};
    candidateSheetData.forEach(function (row) {
        if (row[makeUpColumn] == true) {
            candidateSubjects[row[filteredSubjectCodeColumn]] = {
                isComputerGraded: row[isComputerGradedColumn],
                isManualGraded: row[isManualGradedColumn],
            };
        }
    });

    let openTeacherName = {};
    openData.forEach(function (row) {
        if (row[teacherNameColumn].toString().length > 10) {
            openTeacherName[
                row[openClassNameColumn].toString() +
                    row[openSubjectNameColumn].toString()
            ] = row[teacherNameColumn].toString().split(",")[0].slice(7);
        } else {
            openTeacherName[
                row[openClassNameColumn].toString() +
                    row[openSubjectNameColumn].toString()
            ] = row[teacherNameColumn].toString().slice(6);
        }
    });

    // 清除「排入考程的補考名單」內容
    initialize();

    let nameList = [];
    unfilteredData.forEach(function (row) {
        if (Object.keys(candidateSubjects).includes(row[subjectCodeColumn])) {
            // 科別	年級	班級代碼	班級	座號	學號	姓名	科目名稱	節次	試場	小袋序號	小袋人數	大袋序號	大袋人數	班級人數	時間	電腦	人工	任課老師
            let tmp = [
                getDepartmentName(row[classNameColumn]), // 科別
                getGrade(row[classNameColumn]), // 年級
                getClassCode(row[classNameColumn]), // 班級代碼
                row[classNameColumn], // 班級
                row[seatNumberColumn], // 座號
                row[studentIdColumn], // 學號
                row[studentNameColumn], // 姓名
                row[subjectNameColumn], // 科目名稱
                (row[8] = 0), // 節次預設為0
                (row[9] = 0), // 試場預設為0
                "", // 小袋序號
                "", // 小袋人數
                "", // 大袋序號
                "", // 大袋人數
                "", // 班級人數
                "", // 時間
                candidateSubjects[row[subjectCodeColumn]]["isComputerGraded"]
                    ? "☑"
                    : "☐", // 電腦
                candidateSubjects[row[subjectCodeColumn]]["isManualGraded"]
                    ? "☑"
                    : "☐", //人工
                openTeacherName[
                    row[classNameColumn].toString() +
                        row[subjectNameColumn].toString()
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
