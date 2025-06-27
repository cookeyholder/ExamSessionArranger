/**
 * 註冊組補考名單工作表的課程代碼補完。
 * @returns {void}
 */
function completeUnfilteredSheetCode() {
    const [unfilteredSheetHeaders, ...unfilteredData] = unfilteredSheet
        .getDataRange()
        .getValues();

    const classNameColumnIndex = unfilteredSheetHeaders.indexOf("班級");
    const subjectCodeAndNameColumnIndex =
        unfilteredSheetHeaders.indexOf("科目");
    const subjectCodeColumnIndex =
        unfilteredSheetHeaders.indexOf("科目代碼補完");
    const subjectNameColumnIndex = unfilteredSheetHeaders.indexOf("科目名稱");

    const groupCodeOfDepartment = {
        301: "21",
        303: "22",
        305: "23",
        306: "23",
        308: "23",
        309: "23",
        311: "25",
        315: "24",
        373: "28",
        374: "21",
    };

    const yearOfGrade = {
        一: parseInt(configs["學年度"]),
        二: parseInt(configs["學年度"]) - 1,
        三: parseInt(configs["學年度"]) - 2,
    };

    const parseSubjectCodeAndName = (subjectCodeAndName) =>
        subjectCodeAndName.toString().split(".");

    const getGradeFromClassName = (className) =>
        className.toString().slice(2, 3);

    const getDepartmentCodeFromSubject = (codeString) => codeString.slice(0, 3);

    const buildLongCode = (codeString, schoolCode) =>
        codeString.slice(0, 3) +
        schoolCode +
        codeString.slice(3, 9) +
        "0" +
        codeString.slice(9);

    const buildShortCode = (codeString, className) => {
        const grade = getGradeFromClassName(className);
        const departmentCode = getDepartmentCodeFromSubject(codeString);

        return (
            yearOfGrade[grade] +
            "553401V" +
            groupCodeOfDepartment[departmentCode] +
            codeString.slice(0, 3) +
            "0" +
            codeString.slice(3)
        );
    };

    const completeCode = (codeString, className) =>
        codeString.length === 16
            ? buildLongCode(codeString, configs["學校代碼"])
            : buildShortCode(codeString, className);

    const processRow = (row) => {
        const [codeString, nameString] = parseSubjectCodeAndName(
            row[subjectCodeAndNameColumnIndex]
        );
        const completedCode = completeCode(
            codeString,
            row[classNameColumnIndex]
        );

        return [completedCode, nameString];
    };

    const codeNamePairs = unfilteredData.map(processRow);

    const writeResults = (pairs) => {
        setRangeValues(
            unfilteredSheet.getRange(2, 13, pairs.length, pairs[0].length),
            pairs
        );
        Logger.log(
            "(completeUnfilteredSheetCode) 註冊組補考名單工作表的課程代碼補完成功！"
        );
    };

    const handleError = () => {
        Logger.log(
            "(completeUnfilteredSheetCode) 註冊組補考名單工作表的課程代碼補完失敗！"
        );
        SpreadsheetApp.getUi().alert(
            "註冊組補考名單工作表的課程代碼補完失敗！"
        );
    };

    codeNamePairs.length === unfilteredData.length
        ? writeResults(codeNamePairs)
        : handleError();
}

function open_code_complete() {
    const [openSheetHeaders, ...openData] = openSheet
        .getDataRange()
        .getValues();

    const classNameColumnIndex = openSheetHeaders.indexOf("班級名稱");
    const subjectCodeColumnIndex = openSheetHeaders.indexOf("科目代碼");
    const complete_column = openSheetHeaders.indexOf("科目代碼補完");
    const subjectNameColumnIndex = openSheetHeaders.indexOf("科目名稱");

    const department_to_group = {
        301: "21",
        303: "22",
        305: "23",
        306: "23",
        308: "23",
        309: "23",
        311: "25",
        315: "24",
        373: "28",
        374: "21",
    };

    const yearOfGrade = {
        一: parseInt(configs["學年度"]),
        二: parseInt(configs["學年度"]) - 1,
        三: parseInt(configs["學年度"]) - 2,
    };

    let modified_data = [];
    openData.forEach(function (row) {
        let tmp = row[subjectCodeColumnIndex];
        if (row[subjectCodeColumnIndex].length == 16) {
            row[complete_column] =
                tmp.slice(0, 3) +
                "553401" +
                tmp.slice(3, 9) +
                "0" +
                tmp.slice(9);
        } else {
            row[complete_column] =
                yearOfGrade[row[classNameColumnIndex].toString().slice(2, 3)] +
                "553401V" +
                department_to_group[tmp.slice(0, 3)] +
                tmp.slice(0, 3) +
                "0" +
                tmp.slice(3);
        }

        modified_data.push(row);
    });

    if (modified_data.length == openData.length) {
        setRangeValues(
            openSheet.getRange(
                2,
                1,
                modified_data.length,
                modified_data[0].length
            ),
            modified_data
        );
    } else {
        Logger.log("開課資料課程代碼補完失敗！");
        SpreadsheetApp.getUi().alert("開課資料課程代碼補完失敗！");
    }
}
