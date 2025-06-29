/**
 * 取得共同科目要安排的節次資料
 * @returns {Object} 共同科目要安排的節次資料
 */
function getCommonSubjectSessions() {
    const commonSessions = prearrangedSheet
        .getRange(3, 1, 20, 2)
        .getValues()
        .reduce(function (acc, row) {
            if (row[0] && row[1]) {
                acc[row[0]] = row[1];
            }
            return acc;
        }, {});

    Logger.log(
        "(getCommonSubjectSessions) 共同(預排)科目要安排的節次資料: " +
            JSON.stringify(commonSessions)
    );

    return commonSessions;
}

/**
 * 安排共同科目的節次
 * @returns {void}
 */
function arrangeCommonSubjectSessions() {
    const commonSessions = getCommonSubjectSessions();

    const [filteredHeaders, ...filteredData] = filteredSheet
        .getDataRange()
        .getValues();
    const subjectNameColumn = filteredHeaders.indexOf("科目名稱");
    const sessionColumn = filteredHeaders.indexOf("節次");

    const modifiedData = filteredData.map(function (row) {
        if (commonSessions[row[subjectNameColumn]] == null) {
            return row;
        } else {
            row[sessionColumn] = commonSessions[row[subjectNameColumn]];
            return row;
        }
    });

    if (modifiedData.length === filteredData.length) {
        setRangeValues(
            filteredSheet.getRange(
                2,
                1,
                modifiedData.length,
                modifiedData[0].length
            ),
            modifiedData
        );
    } else {
        Logger.log(
            "(arrangeCommonSubjectSessions)安排共同科目節次時，合併後的資料筆數和原有的筆數不同！"
        );
        SpreadsheetApp.getUi().alert(
            "安排共同科目節次時，合併後的資料筆數和原有的筆數不同！"
        );
    }
}

/**
 * 檢查科別年級是否已存在於節次中
 * @param {Object} session - 節次物件
 * @param {string} departmentGrade - 科別年級組合
 * @returns {boolean} 是否有重複
 */
const hasDepartmentGradeDuplicate = (session, departmentGrade) => {
    const stats = session.departmentGradeStatistics || {};
    return Object.keys(stats).includes(departmentGrade);
};

/**
 * 檢查該節是否有足夠名額
 * @param {Object} session - 節次物件
 * @param {number} additionalStudents - 要加入的學生數
 * @param {number} maxStudents - 最大學生數
 * @returns {boolean} 是否有足夠名額
 */
const hasSessionQuota = (session, additionalStudents, maxStudents) => {
    return additionalStudents + session.population <= maxStudents;
};

/**
 * 檢查科別年級科目組合是否可安排到指定節次
 * @param {Object} session - 節次物件
 * @param {Array} dgsItem - 科別年級科目資料 [key, count]
 * @param {number} maxStudents - 最大學生數
 * @returns {boolean} 是否可安排
 */
const canScheduleToSession = (session, dgsItem, maxStudents) => {
    const departmentGrade = dgsItem[0].slice(0, dgsItem[0].indexOf("_"));
    const studentCount = dgsItem[1];

    return (
        !hasDepartmentGradeDuplicate(session, departmentGrade) &&
        hasSessionQuota(session, studentCount, maxStudents) &&
        session.population < maxStudents
    );
};

/**
 * 將學生安排到指定節次
 * @param {Array} filteredData - 過濾後的學生資料
 * @param {Object} session - 節次物件
 * @param {Array} dgsItem - 科別年級科目資料
 * @param {number} sessionIndex - 節次索引
 * @param {number} sessionColumn - 節次欄位索引
 * @returns {Array} 更新後的學生資料
 */
const assignStudentsToSession = (
    filteredData,
    session,
    dgsItem,
    sessionIndex,
    sessionColumn
) => {
    const targetKey = dgsItem[0];

    return filteredData.map((row) => {
        const departmentColumn = 0;
        const gradeColumn = 1;
        const subjectNameColumn = 7;
        const sessionColumnIndexInFilteredData = 8; // 補考名單中「節次」的欄位順序
        const studentKey =
            row[departmentColumn] +
            row[gradeColumn] +
            "_" +
            row[subjectNameColumn];

        if (
            studentKey === targetKey &&
            row[sessionColumnIndexInFilteredData] === 0
        ) {
            row[sessionColumn] = sessionIndex;
            session.students.push(row);
        }
        return row;
    });
};

/**
 * 處理單一節次的學生安排
 * @param {Array} filteredData - 過濾後的學生資料
 * @param {Array} sessions - 所有節次
 * @param {Array} dgs - 科別年級科目統計
 * @param {number} sessionIndex - 節次索引
 * @param {number} maxStudents - 最大學生數
 * @param {number} sessionColumn - 節次欄位索引
 * @returns {Array} 更新後的學生資料
 */
const processSessionScheduling = (
    filteredData,
    sessions,
    dgs,
    sessionIndex,
    maxStudents,
    sessionColumn
) => {
    const session = sessions[sessionIndex];
    let updatedData = filteredData;

    for (const dgsItem of dgs) {
        if (session.population >= maxStudents) {
            Logger.log(
                `(processSessionScheduling) 第${sessionIndex}節已達人數上限。`
            );
            Logger.log(
                `(processSessionScheduling) 學生數為： ${session.population}`
            );
            Logger.log(
                `(processSessionScheduling) 每節的最大學生數為： ${maxStudents}`
            );
            break;
        }

        if (canScheduleToSession(session, dgsItem, maxStudents)) {
            updatedData = assignStudentsToSession(
                filteredData,
                session,
                dgsItem,
                sessionIndex,
                sessionColumn
            );
        }
    }

    return updatedData;
};

/**
 * 收集所有節次的學生資料
 * @param {Array} sessions - 所有節次
 * @param {number} maxSessionNumber - 最大節次數
 * @returns {Array} 合併後的學生資料
 */
const collectAllSessionStudents = (sessions, maxSessionNumber) => {
    return Array.from({ length: maxSessionNumber + 1 }, (_, i) => i + 1).reduce(
        (acc, sessionIndex) => {
            const session = sessions[sessionIndex];
            if (session) {
                Logger.log(
                    `(collectAllSessionStudents) 第${sessionIndex}節應考學生數: ${session.population}位。`
                );
                return acc.concat(session.students);
            } else {
                Logger.log(
                    `(collectAllSessionStudents) 第${sessionIndex}節沒有學生資料。`
                );
                return acc;
            }
        },
        []
    );
};

/**
 * 安排專業科目(非共同科目)的節次
 * @returns {void}
 */
function arrangeProfessionsSession() {
    const [filteredHeaders, ...filteredData] = filteredSheet
        .getDataRange()
        .getValues();
    const sessionColumn = filteredHeaders.indexOf("節次");

    const MAX_SESSION_NUMBER = parseInt(configs["節數上限"]);
    const NUMBER_OF_CLASSROOMS = parseInt(configs["試場數量"]);
    const MAX_STUDENTS_PER_CLASSROOM = parseInt(configs["每間試場人數上限"]);
    const MAX_STUDENTS_PER_SESSION =
        NUMBER_OF_CLASSROOMS * MAX_STUDENTS_PER_CLASSROOM;

    const dgs = getDepartmentGradeSubjectPopulation();
    const sessions = getSessionStatistics();

    // 處理所有節次的學生安排
    let updatedData = filteredData;
    for (let i = 1; i <= MAX_SESSION_NUMBER + 1; i++) {
        updatedData = processSessionScheduling(
            updatedData,
            sessions,
            dgs,
            i,
            MAX_STUDENTS_PER_SESSION,
            sessionColumn
        );
    }

    // 收集所有節次的學生資料
    const modifiedData = collectAllSessionStudents(
        sessions,
        MAX_SESSION_NUMBER
    );

    Logger.log(
        `(arrangeProfessionsSession) 原始資料筆數：${filteredData.length} `
    );
    Logger.log(
        `(arrangeProfessionsSession) 安排好的專業科目節次資料筆數：${modifiedData.length} `
    );

    if (modifiedData.length === filteredData.length) {
        setRangeValues(
            filteredSheet.getRange(
                2,
                1,
                modifiedData.length,
                modifiedData[0].length
            ),
            modifiedData
        );
    } else {
        Logger.log(
            `(arrangeProfessionsSession) 無法將所有人排入 ${MAX_SESSION_NUMBER} 節，以下的「科別-年級」組合需補考的科目超過 ${MAX_SESSION_NUMBER} 科：${findExcessiveValuesInDgs(
                dgs,
                MAX_SESSION_NUMBER
            )}`
        );
        SpreadsheetApp.getUi().alert(
            `無法將所有人排入 ${MAX_SESSION_NUMBER} 節，以下的「科別-年級」組合需補考的科目超過 ${MAX_SESSION_NUMBER} 科：${findExcessiveValuesInDgs(
                dgs,
                MAX_SESSION_NUMBER
            )}`
        );
    }
}

/**
 * 找出 dgs 中人數超過節次上限的科別年級科目組合
 * @param {Array} dgs - 科別年級科目統計資料
 * @param {number} maxSessionNumber - 最大節次數
 * @returns {string} 超過上限的科別年級科目組合及其次數的描述
 */
function findExcessiveValuesInDgs(dgs, maxSessionNumber) {
    const departmentGradeCounts = dgs
        .map((dgsItem) => dgsItem[0].split("_")[0])
        .reduce((acc, departmentGrade) => {
            acc[departmentGrade] = (acc[departmentGrade] || 0) + 1;
            return acc;
        }, {});

    const excessiveItems = Object.keys(departmentGradeCounts)
        .filter(
            (departmentGrade) =>
                departmentGradeCounts[departmentGrade] > maxSessionNumber
        )
        .map(
            (departmentGrade) =>
                `${departmentGrade}年級(${departmentGradeCounts[departmentGrade]}科)`
        );

    return excessiveItems.length > 0 ? excessiveItems.join(", ") : "無";
}
