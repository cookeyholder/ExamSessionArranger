function generate_sessionRecordSheet() {
    sortBySessionThenClassroom();

    const [headers, ...data] = filteredSheet.getDataRange().getValues();

    const session_column = headers.indexOf("節次");
    const classroom_column = headers.indexOf("試場");
    const time_column = headers.indexOf("時間");
    const classNameColumn = headers.indexOf("班級");
    const studentIdColumn = headers.indexOf("學號");
    const name_column = headers.indexOf("姓名");
    const subject_column = headers.indexOf("科目名稱");
    const class_population_column = headers.indexOf("班級人數");

    // 刪除多餘的欄和列
    sessionRecordSheet.clear();
    if (sessionRecordSheet.getMaxRows() > 5) {
        sessionRecordSheet.deleteRows(2, sessionRecordSheet.getMaxRows() - 5);
    }

    let modified_data = [
        [
            "A表：112學年度第1學期補考簽到及違規記錄表    　 　　　                                 監考教師簽名：　　　　　　　　　",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ],
        [
            "節次",
            "試場",
            "時間",
            "班級",
            "學號",
            "姓名",
            "科目名稱",
            "班級人數",
            "考生到考簽名",
            "違規記錄(打V)",
            "",
            "其他違規\n請簡述",
        ],
        ["", "", "", "", "", "", "", "", "", "未帶有照證件", "服儀不整", ""],
    ];

    data.forEach(function (row) {
        modified_data.push([
            row[session_column],
            row[classroom_column],
            row[time_column],
            row[classNameColumn],
            row[studentIdColumn],
            row[name_column],
            row[subject_column],
            row[class_population_column],
            "",
            "",
            "",
            "",
        ]);
    });

    setRangeValues(
        sessionRecordSheet.getRange(
            1,
            1,
            modified_data.length,
            modified_data[0].length
        ),
        modified_data
    );

    // 設定格式美化表格
    sessionRecordSheet
        .getRange("A1:L1")
        .mergeAcross()
        .setVerticalAlignment("bottom")
        .setFontSize(14)
        .setFontWeight("bold");
    sessionRecordSheet.getRange("J2:K2").mergeAcross();
    sessionRecordSheet
        .getRange("A2:A3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    sessionRecordSheet
        .getRange("B2:B3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    sessionRecordSheet
        .getRange("C2:C3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    sessionRecordSheet
        .getRange("D2:D3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    sessionRecordSheet
        .getRange("E2:E3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    sessionRecordSheet
        .getRange("F2:F3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    sessionRecordSheet
        .getRange("G2:G3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    sessionRecordSheet
        .getRange("H2:H3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    sessionRecordSheet
        .getRange("I2:I3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    sessionRecordSheet
        .getRange("L2:L3")
        .mergeVertically()
        .setVerticalAlignment("middle");

    sessionRecordSheet
        .getRange(2, 1, modified_data.length + 2, modified_data[0].length)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
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
