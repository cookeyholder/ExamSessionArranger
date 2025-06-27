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

    let codeNamePairs = unfilteredData.map((row) => {
        let [codeString, nameString] = row[subjectCodeAndNameColumnIndex]
            .toString()
            .split(".");

        if (codeString.length == 16) {
            code =
                codeString.slice(0, 3) +
                configs["學校代碼"] +
                codeString.slice(3, 9) +
                "0" +
                codeString.slice(9);
        } else {
            code =
                yearOfGrade[row[classNameColumnIndex].toString().slice(2, 3)] +
                "553401V" +
                groupCodeOfDepartment[codeString.slice(0, 3)] +
                codeString.slice(0, 3) +
                "0" +
                codeString.slice(3);
        }

        return [code, nameString];
    });

    if (codeNamePairs.length == unfilteredData.length) {
        setRangeValues(
            unfilteredSheet.getRange(
                2,
                13,
                codeNamePairs.length,
                codeNamePairs[0].length
            ),
            codeNamePairs
        );
        Logger.log(
            "(completeUnfilteredSheetCode) 註冊組補考名單工作表的課程代碼補完成功！"
        );
    } else {
        Logger.log(
            "(completeUnfilteredSheetCode) 註冊組補考名單工作表的課程代碼補完失敗！"
        );
        SpreadsheetApp.getUi().alert(
            "註冊組補考名單工作表的課程代碼補完失敗！"
        );
    }
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
