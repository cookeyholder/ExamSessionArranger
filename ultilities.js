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
        Logger.log(
            `(setRangeValues) 欲寫入的範圍大小: ${range.getNumRows()}列 x ${range.getNumColumns()}行`
        );
        Logger.log(
            `(setRangeValues) 欲寫入的 data 大小: ${data.length}列 x ${data[0].length}行`
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

/**
 * 統計各科別年級、各班級、科目的應考人數
 * @returns {Object} 包含科別、年級、班級和科目的統計數據
 */
function getDepartmentGradeSubjectPopulation() {
    const [filteredHeaders, ...filteredData] = filteredSheet
        .getDataRange()
        .getValues();

    const departmentColumn = filteredHeaders.indexOf("科別");
    const gradeColumn = filteredHeaders.indexOf("年級");
    const subjectNameColumn = filteredHeaders.indexOf("科目名稱");

    const createStatisticsKey = (row) =>
        row[departmentColumn] + row[gradeColumn] + "_" + row[subjectNameColumn];

    const updateStatistics = (statistics, row) => {
        const key = createStatisticsKey(row);
        return {
            ...statistics,
            [key]: (statistics[key] || 0) + 1,
        };
    };

    return Object.entries(filteredData.reduce(updateStatistics, {})).sort(
        descending_sorting
    );
}

/**
 * 建立教室物件的工廠函數
 * @param {Array} students - 學生資料陣列
 * @returns {Object} 教室物件
 */
const createClassroom = (students = []) => ({
    students,
    get population() {
        return this.students.length;
    },
    get classSubjectStatistics() {
        return this.students.reduce((statistics, row) => {
            const key = row[3] + "_" + row[7]; // 班級 + _ + 科目
            return {
                ...statistics,
                [key]: (statistics[key] || 0) + 1,
            };
        }, {});
    },
});

/**
 * 建立節次物件的工廠函數
 * @param {Array} students - 學生資料陣列
 * @param {number} maxClassroomNumber - 最大試場數量
 * @returns {Object} 節次物件
 */
const createSession = (students = [], maxClassroomNumber) => ({
    classrooms: Array.from({ length: maxClassroomNumber + 1 }, () =>
        createClassroom()
    ),
    students,
    get population() {
        return this.students.length;
    },
    get departmentGradeStatistics() {
        return this.students.reduce((statistics, row) => {
            const key = row[0] + row[1]; // 科別 + 年級
            return {
                ...statistics,
                [key]: (statistics[key] || 0) + 1,
            };
        }, {});
    },
    get departmentClassSubjectStatistics() {
        return this.students.reduce((statistics, row) => {
            const key = row[3] + row[7]; // 班級 + 科目
            return {
                ...statistics,
                [key]: (statistics[key] || 0) + 1,
            };
        }, {});
    },
});

/**
 * 取得節次統計資料
 * @returns {Array} 包含所有節次統計資料的陣列
 */
function getSessionStatistics() {
    const [headers, ...data] = filteredSheet.getDataRange().getValues();
    const sessionColumn = headers.indexOf("節次");
    const MAX_SESSION_NUMBER = parseInt(configs["節數上限"]);
    const MAX_CLASSROOM_NUMBER = parseInt(configs["試場數量"]);

    // 建立空的節次陣列
    const sessions = Array.from({ length: MAX_SESSION_NUMBER + 2 }, () =>
        createSession([], MAX_CLASSROOM_NUMBER)
    );

    // 將學生資料按節次分組
    const studentsBySession = data.reduce((acc, row) => {
        const sessionIndex = row[sessionColumn];
        if (!acc[sessionIndex]) {
            acc[sessionIndex] = [];
        }
        acc[sessionIndex].push(row);
        return acc;
    }, {});

    // 為每個節次建立包含學生資料的節次物件
    return sessions.map((_, index) =>
        createSession(studentsBySession[index] || [], MAX_CLASSROOM_NUMBER)
    );
}
