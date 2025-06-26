function unfilted_code_complete() {
    const [unfilteredSheetHeaders, ...unfilteredData] = unfilteredSheet
        .getDataRange()
        .getValues();

    const classNameColumn = unfilteredSheetHeaders.indexOf("班級");
    const subject_column = unfilteredSheetHeaders.indexOf("科目");
    const subjectCodeColumn = unfilteredSheetHeaders.indexOf("科目代碼補完");
    const subjectNameColumn = unfilteredSheetHeaders.indexOf("科目名稱");

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

    const grade_to_year = {
        一: parseInt(parametersSheet.getRange("B2").getValue()),
        二: parseInt(parametersSheet.getRange("B2").getValue()) - 1,
        三: parseInt(parametersSheet.getRange("B2").getValue()) - 2,
    };

    let modified_data = [];
    unfilteredData.forEach(function (row) {
        let tmp = row[subject_column].toString().split(".")[0];
        if (tmp.length == 16) {
            row[subjectCodeColumn] =
                tmp.slice(0, 3) +
                "553401" +
                tmp.slice(3, 9) +
                "0" +
                tmp.slice(9);
        } else {
            row[subjectCodeColumn] =
                grade_to_year[row[classNameColumn].toString().slice(2, 3)] +
                "553401V" +
                department_to_group[tmp.slice(0, 3)] +
                tmp.slice(0, 3) +
                "0" +
                tmp.slice(3);
        }

        row[subjectNameColumn] = row[subject_column].toString().split(".")[1];
        modified_data.push(row);
    });

    if (modified_data.length == unfilteredData.length) {
        set_range_values(
            unfilteredSheet.getRange(
                2,
                1,
                modified_data.length,
                modified_data[0].length
            ),
            modified_data
        );
    } else {
        Logger.log("課程代碼補完失敗！");
        SpreadsheetApp.getUi().alert("課程代碼補完失敗！");
    }
}

function open_code_complete() {
    const [openSheetHeaders, ...openData] = openSheet
        .getDataRange()
        .getValues();

    const classNameColumn = openSheetHeaders.indexOf("班級名稱");
    const subjectCodeColumn = openSheetHeaders.indexOf("科目代碼");
    const complete_column = openSheetHeaders.indexOf("科目代碼補完");
    const subjectNameColumn = openSheetHeaders.indexOf("科目名稱");

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

    const grade_to_year = {
        一: parseInt(parametersSheet.getRange("B2").getValue()),
        二: parseInt(parametersSheet.getRange("B2").getValue()) - 1,
        三: parseInt(parametersSheet.getRange("B2").getValue()) - 2,
    };

    let modified_data = [];
    openData.forEach(function (row) {
        let tmp = row[subjectCodeColumn];
        if (row[subjectCodeColumn].length == 16) {
            row[complete_column] =
                tmp.slice(0, 3) +
                "553401" +
                tmp.slice(3, 9) +
                "0" +
                tmp.slice(9);
        } else {
            row[complete_column] =
                grade_to_year[row[classNameColumn].toString().slice(2, 3)] +
                "553401V" +
                department_to_group[tmp.slice(0, 3)] +
                tmp.slice(0, 3) +
                "0" +
                tmp.slice(3);
        }

        modified_data.push(row);
    });

    if (modified_data.length == openData.length) {
        set_range_values(
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
