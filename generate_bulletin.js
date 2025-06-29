function generate_bulletin() {
    sortByClassName();

    const [headers, ...data] = filteredSheet.getDataRange().getValues();

    const classNameColumn = headers.indexOf("班級");
    const studentIdColumn = headers.indexOf("學號");
    const name_column = headers.indexOf("姓名");
    const subject_column = headers.indexOf("科目名稱");
    const session_column = headers.indexOf("節次");
    const classroom_column = headers.indexOf("試場");

    // 刪除多餘的欄和列
    bulletinSheet.clear();
    if (bulletinSheet.getMaxRows() > 5) {
        bulletinSheet.deleteRows(2, bulletinSheet.getMaxRows() - 5);
    }

    let modified_data = [["班級", "學號", "姓名", "科目", "節次", "試場"]];
    data.forEach(function (row) {
        let repeat_times = 0;
        let masked_name = "";

        if (row[name_column].length == 2) {
            masked_name = row[name_column].toString().slice(0, 1) + "〇";
        } else {
            repeat_times = row[name_column].length - 2;
            masked_name =
                row[name_column].toString().slice(0, 1) +
                "〇".repeat(repeat_times) +
                row[name_column].toString().slice(-1);
        }

        let tmp = [
            row[classNameColumn],
            row[studentIdColumn],
            masked_name,
            row[subject_column],
            row[session_column],
            row[classroom_column],
        ];

        modified_data.push(tmp);
    });

    setRangeValues(
        bulletinSheet.getRange(
            2,
            1,
            modified_data.length,
            modified_data[0].length
        ),
        modified_data
    );
    sortBySessionThenClassroom();
    prettier();
}

function prettier() {
    const school_year = parametersSheet.getRange("B2").getValue();
    const semester = parametersSheet.getRange("B3").getValue();

    bulletinSheet.getRange("A1:F1").mergeAcross();
    bulletinSheet
        .getRange("A1")
        .setValue(
            "高雄高工" + school_year + "學年度第" + semester + "學期補考名單"
        );
    bulletinSheet.getRange("A1").setFontSize(20);
    bulletinSheet
        .getRange(
            1,
            1,
            bulletinSheet.getMaxRows(),
            bulletinSheet.getMaxColumns()
        )
        .setHorizontalAlignment("center");
    bulletinSheet.setFrozenRows(2);
    bulletinSheet.getRange("A2:F").createFilter();
    bulletinSheet
        .getRange("A2:F")
        .setBorder(
            true,
            true,
            true,
            true,
            true,
            true,
            "#000000",
            SpreadsheetApp.BorderStyle.SOLID
        );
}
