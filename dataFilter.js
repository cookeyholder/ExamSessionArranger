function getFilteredData() {
    // 清除「排入考程的補考名單」工作表的內容
    initialize();

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
    for (const row of candidateSheetData) {
        const needsMakeUpExam = row[makeUpColumn] === true;

        if (needsMakeUpExam) {
            const subjectCode = row[filteredSubjectCodeColumn];
            candidateSubjects[subjectCode] = {
                isComputerGraded: row[isComputerGradedColumn],
                isManualGraded: row[isManualGradedColumn],
            };
        }
    }

    // 建立班級科目與任課教師的對應表
    const createClassSubjectKey = (className, subjectName) =>
        className.toString() + subjectName.toString();

    const extractTeacherName = (teacherNameText) => {
        const teacherName = teacherNameText.toString();
        return teacherName.length > 10
            ? teacherName.split(",")[0].slice(7)
            : teacherName.slice(6);
    };

    const openTeacherName = Object.fromEntries(
        openData.map((row) => [
            createClassSubjectKey(
                row[openClassNameColumn],
                row[openSubjectNameColumn]
            ),
            extractTeacherName(row[teacherNameColumn]),
        ])
    );

    const createStudentRecord = (row) => {
        const subjectCode = row[subjectCodeColumn];
        const subjectInfo = candidateSubjects[subjectCode];

        return [
            getDepartmentName(row[classNameColumn]), // 科別
            getGrade(row[classNameColumn]), // 年級
            getClassCode(row[classNameColumn]), // 班級代碼
            row[classNameColumn], // 班級
            row[seatNumberColumn], // 座號
            row[studentIdColumn], // 學號
            row[studentNameColumn], // 姓名
            row[subjectNameColumn], // 科目名稱
            0, // 節次預設為0
            0, // 試場預設為0
            "", // 小袋序號
            "", // 小袋人數
            "", // 大袋序號
            "", // 大袋人數
            "", // 班級人數
            "", // 時間
            subjectInfo.isComputerGraded ? "☑" : "☐", // 電腦
            subjectInfo.isManualGraded ? "☑" : "☐", // 人工
            openTeacherName[
                createClassSubjectKey(
                    row[classNameColumn],
                    row[subjectNameColumn]
                )
            ], // 任課老師
        ];
    };

    const isEligibleForMakeUp = (row) =>
        candidateSubjects.hasOwnProperty(row[subjectCodeColumn]);

    const nameList = unfilteredData
        .filter(isEligibleForMakeUp)
        .map(createStudentRecord);

    filteredSheet
        .getRange(2, 1, nameList.length, nameList[0].length)
        .setNumberFormat("@STRING@") // 改成純文字格式，以免 0 開頭的學號被去掉前面的 0，造成位數錯誤
        .setValues(nameList);
}
