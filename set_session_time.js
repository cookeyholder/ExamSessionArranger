function set_session_time() {
    const [headers, ...data] = filteredSheet.getDataRange().getValues();
    const session_column = headers.indexOf("節次");
    const time_column = headers.indexOf("時間");

    const timeTableSheet = ss.getSheetByName("節次時間表");
    const [time_headers, ...time_data] = timeTableSheet
        .getDataRange()
        .getValues();

    const time_table = {};
    time_data.forEach(function (time_row) {
        time_table[time_row[0]] = time_row[1];
    });

    let modified_data = data.map(function (row) {
        row[time_column] = time_table[row[session_column]];
        return row;
    });

    if (modified_data.length == data.length) {
        setRangeValues(
            filteredSheet.getRange(
                2,
                1,
                modified_data.length,
                modified_data[0].length
            ),
            modified_data
        );
    } else {
        Logger.log("寫入節次時間失敗！");
        SpreadsheetApp.getUi().alert("寫入節次時間失敗！");
    }
}
