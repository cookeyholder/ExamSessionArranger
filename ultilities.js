/**
 * 取得參數設定表中的配置。
 * 這個函數會讀取名為 "參數設定" 的工作表，並將第一列的值作為鍵，第二列的值作為對應的值，
 * 返回一個包含所有配置的物件。
 *
 * @returns {Object} 包含參數設定的物件。
 */
function getConfigs() {
    let configs = {};
    parametersSheet
        .getDataRange()
        .getValues()
        .reduce(function (acc, row) {
            if (row[0] && row[1]) {
                acc[row[0]] = row[1];
            }
            return acc;
        }, configs);

    Logger.log("(getConfigs) configs: " + JSON.stringify(configs));
    return configs;
}

function calculateElapsedTimeInSeconds(startTime) {
    let stopTime = new Date();
    let newRuntime = Number(stopTime) - Number(startTime);
    return Math.ceil(newRuntime / 1000);
}

function count_time_consume(runner) {
    let start_time = new Date();
    runner();
    let end_time = new Date();
    let runtime = Math.ceil(Number(end_time) - Number(start_time)) / 1000;
    return runtime;
}

/** * 設定指定範圍的值。
 * 如果範圍的行數與數據的行數不匹配，則會顯示一個警告對話框。
 * @param {Range} range - 要設置值的範圍。
 * @param {Array} data - 要設置的數據，應為二維數組。
 */
function setRangeValues(range, data) {
    if (
        range.getNumRows() == data.length &&
        range.getNumColumns() == data[0].length
    ) {
        range.setValues(data);
        return true;
    } else {
        Logger.log("(setRangeValues) 欲寫入的範圍大小和 data 的大小不一致！");
        // 顯示 range 的大小
        Logger.log(
            "(setRangeValues) 欲寫入的範圍大小: " +
                range.getNumRows() +
                "列 x " +
                range.getNumColumns() +
                "行"
        );
        Logger.log(
            "(setRangeValues) 欲寫入的 data 大小: " +
                data.length +
                "列 x " +
                data[0].length +
                "行"
        );
        SpreadsheetApp.getUi().alert("欲寫入的範圍大小和 data 的大小不一致！");
        return false;
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

    setRangeValues(data_range, data_values);
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

    setRangeValues(data_range, data_values);
}

function descending_population(a, b) {
    if (a[1] === b[1]) {
        return 0;
    } else {
        return a[1] < b[1] ? 1 : -1;
    }
}

/**
 * 統計各科別年級、各班級、科目的應考人數
 * @returns {Object} 包含科別、年級、班級和科目的統計數據
 */
function getDepartmentGradeSubjectCounts() {
	const [filteredHeaders, ...filteredData] = filteredSheet
		.getDataRange()
		.getValues();

	const departmentColumn = filteredHeaders.indexOf('科別');
	const gradeColumn = filteredHeaders.indexOf('年級');
	const subjectNameColumn = filteredHeaders.indexOf('科目名稱');

	const createStatisticsKey = row => 
		row[departmentColumn] + row[gradeColumn] + '_' + row[subjectNameColumn];

	const updateStatistics = (statistics, row) => {
		const key = createStatisticsKey(row);
		return {
			...statistics,
			[key]: (statistics[key] || 0) + 1
		};
	};

	return filteredData.reduce(updateStatistics, {});
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
    const MAX_CLASSROOM_NUMBER = parseInt(configs["試場數量"]);
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
    const MAX_SESSION_NUMBER = parserseInt(configs["節數上限"]);

    const sessions = [];
    for (let i = 0; i < MAX_SESSION_NUMBER + 2; i++) {
        sessions.push(create_session());
    }

    for (row of data) {
        sessions[row[session_column]].students.push(row);
    }

    return sessions;
}
    }

    return sessions;
}
