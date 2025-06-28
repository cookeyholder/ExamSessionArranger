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

    const [headers, ...data] = filteredSheet.getDataRange().getValues();
    const subjectNameColumn = headers.indexOf("科目名稱");
    const sessionColumn = headers.indexOf("節次");

    const modifiedData = data.map(function (row) {
        if (commonSessions[row[subjectNameColumn]] == null) {
            return row;
        } else {
            row[sessionColumn] = commonSessions[row[subjectNameColumn]];
            return row;
        }
    });

    if (modifiedData.length === data.length) {
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

function descending_sorting(a, b) {
    if (a[1] === b[1]) {
        return 0;
    } else {
        return a[1] < b[1] ? 1 : -1;
    }
}

/**
 * 安排專業科目(非共同科目)的節次
 * @returns {void}
 */
function arrangeProfessionsSession() {
    const [headers, ...data] = filteredSheet.getDataRange().getValues();
    const session_column = headers.indexOf("節次");

    const MAX_SESSION_NUMBER = parseInt(configs["節數上限"]);
    const MAX_SESSION_STUDENTS = 0.9 * parseInt(configs["每間試場人數上限"]); // 每節的最大學生數的 9 成

    const dgs = Object.entries(getDepartmentGradeSubjectCounts()).sort(
        descending_sorting
    );
    const sessions = get_session_statistics();

    for (let i = 1; i < MAX_SESSION_NUMBER + 2; i++) {
        for (let k = 0; k < dgs.length; k++) {
            const department_grade = dgs[k][0].slice(0, dgs[k][0].indexOf("_"));
            const has_duplicate = Object.keys(
                sessions[i].department_grade_statisics
            ).includes(department_grade);
            if (has_duplicate) {
                continue;
            }

            const has_quota =
                dgs[k][1] + sessions[i].population <= MAX_SESSION_STUDENTS;
            if (!has_quota) {
                continue;
            }

            if (sessions[i].population >= MAX_SESSION_STUDENTS) {
                Logger.log("第" + i + "節已達人數上限。");
                Logger.log("學生數為： " + sessions[i].population);
                break;
            }

            if (!has_duplicate && has_quota) {
                data.forEach(function (row) {
                    let key = row[0] + row[1] + "_" + row[7];
                    if (key == dgs[k][0] && row[8] == 0) {
                        row[session_column] = i;
                        sessions[i].students.push(row);
                    }
                });
            }
        }
    }

    let modified_data = [];
    for (let i = 1; i < MAX_SESSION_NUMBER + 2; i++) {
        Logger.log("sessions[" + i + "]: " + sessions[i].population);
        modified_data = modified_data.concat(sessions[i].students);
    }

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
        Logger.log(
            "無法將所有人排入 " +
                MAX_SESSION_NUMBER +
                " 節，請檢查是否有某科年級須補考 " +
                parseInt(MAX_SESSION_NUMBER + 1) +
                " 科以上！"
        );
        SpreadsheetApp.getUi().alert(
            "無法將所有人排入 " +
                MAX_SESSION_NUMBER +
                "節，請檢查是否有某科年級須補考 10 科以上！"
        );
    }
}
