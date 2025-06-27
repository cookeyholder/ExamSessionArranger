function generate_small_bag_data() {
    const [headers, ...data] = filteredSheet.getDataRange().getValues();

    const school_year = parametersSheet.getRange("B2").getValue();
    const semester = parametersSheet.getRange("B3").getValue();

    const small_bag_serial_column = headers.indexOf("小袋序號");
    const session_column = headers.indexOf("節次");
    const time_column = headers.indexOf("時間");
    const classroom_column = headers.indexOf("試場");
    const classNameColumn = headers.indexOf("班級");
    const subjectNameColumn = headers.indexOf("科目名稱");
    const teacherNameColumn = headers.indexOf("任課老師");
    const small_bag_population_column = headers.indexOf("小袋人數");
    const isComputerGradedColumn = headers.indexOf("電腦");
    const isManualGradedColumn = headers.indexOf("人工");

    smallBagSheet.clear();

    // 刪除多餘的欄和列，並設置標題列
    if (smallBagSheet.getMaxRows() > 5) {
        smallBagSheet.deleteRows(2, smallBagSheet.getMaxRows() - 5);
    }

    let small_bags = [
        [
            "學年度",
            "學期",
            "小袋序號",
            "節次",
            "時間",
            "試場",
            "班級",
            "科目名稱",
            "任課老師",
            "小袋人數",
            "電腦",
            "人工",
        ],
    ];
    let already_arranged = [];

    data.forEach(function (row) {
        if (!already_arranged.includes(row[small_bag_serial_column])) {
            let tmp = [
                school_year,
                semester,
                row[small_bag_serial_column],
                row[session_column],
                row[time_column],
                row[classroom_column],
                row[classNameColumn],
                row[subjectNameColumn],
                row[teacherNameColumn],
                row[small_bag_population_column],
                row[isComputerGradedColumn],
                row[isManualGradedColumn],
            ];

            small_bags.push(tmp);
            already_arranged.push(row[small_bag_serial_column]);
        }
    });

    setRangeValues(
        smallBagSheet.getRange(1, 1, small_bags.length, small_bags[0].length),
        small_bags
    );
}
