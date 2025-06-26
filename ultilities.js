function runtime_count_stop(start) {
    let stop = new Date();
    let newRuntime = Number(stop) - Number(start);
    return Math.ceil(newRuntime / 1000);
}

function count_time_consume(runner) {
    let start_time = new Date();
    runner();
    let end_time = new Date();
    let runtime = Math.ceil(Number(end_time) - Number(start_time)) / 1000;
    return runtime;
}

function set_range_values(range, data) {
    if (range.getLastColumn() == data[0].length) {
        range.setValues(data);
    } else {
        SpreadsheetApp.getUi().alert("欲寫入範圍欄數不足！");
    }
}

function getClassCode(cls) {
    // 班級代碼查詢
    // 輸入班級中文名稱(4字)，輸出6碼數字代碼
    // 如輸入「機械二丁」，輸出「301204」。

    const departmentCodes = {
        機械: "301",
        汽車: "303",
        資訊: "305",
        電子: "306",
        電機: "308",
        冷凍: "309",
        建築: "311",
        化工: "315",
        圖傳: "373",
        電圖: "374",
    };

    const classAndGradeCode = {
        甲: "01",
        乙: "02",
        丙: "03",
        丁: "04",
        一: "1",
        二: "2",
        三: "3",
    };

    return (
        departmentCodes[cls.slice(0, 2)] +
        classAndGradeCode[cls.slice(2, 3)] +
        classAndGradeCode[cls.slice(-1)]
    );
}

function getGrade(cls) {
    const grade = {
        一: "1",
        二: "2",
        三: "3",
    };

    return grade[cls.slice(2, 3)];
}

function getDepartmentName(cls) {
    const departments = {
        機械: "機械科",
        汽車: "汽車科",
        資訊: "資訊科",
        電子: "電子科",
        電機: "電機科",
        冷凍: "冷凍空調科",
        建築: "建築科",
        化工: "化工科",
        圖傳: "圖文傳播科",
        電圖: "電腦機械製圖科",
    };

    return departments[cls.slice(0, 2)];
}

function checkShowedBoxes() {
    const data_range = ss
        .getSheetByName("教學組排入考程的科目")
        .getRange("A2:A");
    const data_values = data_range.getValues();
    const num_rows = data_range.getNumRows();
    const num_cols = data_range.getNumColumns();

    for (let i = 0; i < num_rows; i++) {
        if (!ss.isRowHiddenByFilter(i + 1)) {
            for (let j = 0; j < num_cols; j++) {
                data_values[i][j] = true;
            }
        }
    }

    set_range_values(data_range, data_values);
}

function cancelCheckboxes() {
    const data_range = ss
        .getSheetByName("教學組排入考程的科目")
        .getRange("A1:A");
    const data_values = data_range.getValues();
    const num_rows = data_range.getNumRows();
    const num_cols = data_range.getNumColumns();

    for (let i = 0; i < num_rows; i++) {
        if (!ss.isRowHiddenByFilter(i + 1)) {
            for (let j = 0; j < num_cols; j++) {
                data_values[i][j] = false;
            }
        }
    }

    set_range_values(data_range, data_values);
}

function descending_population(a, b) {
    if (a[1] === b[1]) {
        return 0;
    } else {
        return a[1] < b[1] ? 1 : -1;
    }
}

function get_department_grade_statistics_of_array(data) {
    let department_column = 0;
    let grade_column = 1;
    let statistics = {};
    for (row of data) {
        let key = row[department_column] + row[grade_column];

        if (key in statistics) {
            statistics[key] += 1;
        } else {
            statistics[key] = 1;
        }
    }

    return statistics;
}

function get_department_grade_subject_statistics_of_array(data) {
    const department_column = 0;
    const grade_column = 1;
    const subjectNameColumn = 7;

    let statistics = {};
    for (row of data) {
        let key =
            row[department_column] +
            row[grade_column] +
            "_" +
            row[subjectNameColumn];

        if (key in statistics) {
            statistics[key] += 1;
        } else {
            statistics[key] = 1;
        }
    }

    return statistics;
}

function get_department_grade_statistics() {
    // 統計各科別年級、各班級的應考人數

    const [headers, ...data] = filteredSheet.getDataRange().getValues();

    return get_department_grade_statistics_of_array(data);
}

function get_department_grade_subject_statistics() {
    // 統計各科別年級、各班級、科目的應考人數

    const [headers, ...data] = filteredSheet.getDataRange().getValues();

    return get_department_grade_subject_statistics_of_array(data);
}

function create_classroom() {
    return {
        students: [],
        get population() {
            return this.students.length;
        },
        get class_subject_statisics() {
            let statistics = {};
            this.students.forEach(function (row) {
                let key = row[3] + "_" + row[7]; // 班級 + _ + 科目
                if (Object.keys(statistics).includes(key)) {
                    statistics[key] += 1;
                } else {
                    statistics[key] = 1;
                }
            });
            return statistics;
        },
    };
}

function create_session() {
    // session 物件工廠，用來產生下面的 get_session_statistic 函數中，需要建立 9 個 session 物件
    const MAX_CLASSROOM_NUMBER = parametersSheet.getRange("B6").getValue();
    const session = {
        classrooms: [],
        students: [],
        get population() {
            return this.students.length;
        },

        get department_grade_statisics() {
            let statistics = {};
            this.students.forEach(function (row) {
                let key = row[0] + row[1]; // 科別 + 年級
                if (Object.keys(statistics).includes(key)) {
                    statistics[key] += 1;
                } else {
                    statistics[key] = 1;
                }
            });
            return statistics;
        },

        get department_class_subject_statisics() {
            let statistics = {};
            this.students.forEach(function (row) {
                let key = row[3] + row[7]; // 班級 + 科目
                if (Object.keys(statistics).includes(key)) {
                    statistics[key] += 1;
                } else {
                    statistics[key] = 1;
                }
            });
            return statistics;
        },
    };

    for (let j = 0; j < MAX_CLASSROOM_NUMBER + 1; j++) {
        session.classrooms.push(create_classroom());
    }
    return session;
}

function get_session_statistics() {
    const [headers, ...data] = filteredSheet.getDataRange().getValues();
    const session_column = headers.indexOf("節次");
    const MAX_SESSION_NUMBER = parametersSheet.getRange("B5").getValue();

    const sessions = [];
    for (let i = 0; i < MAX_SESSION_NUMBER + 2; i++) {
        sessions.push(create_session());
    }

    for (row of data) {
        sessions[row[session_column]].students.push(row);
    }

    return sessions;
}
