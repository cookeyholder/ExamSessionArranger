/**
 * 對兩個數組進行降序排序, 比較第二個(index 為 1)的元素。
 * @param {Array} a - 第一個數組。
 * @param {Array} b - 第二個數組。
 * @returns {number} - 返回比較結果，負數表示 a 在 b 前，
 *                     正數表示 b 在 a 前，0 表示相等。
 */
function descending_sorting(a, b) {
    if (a[1] === b[1]) {
        return 0;
    } else {
        return a[1] < b[1] ? 1 : -1;
    }
}

function arrangeClassroom() {
    const [headers, ...data] = filteredSheet.getDataRange().getValues();
    const classNameColumn = headers.indexOf("班級");
    const subject_column = headers.indexOf("科目名稱");
    const classroom_column = headers.indexOf("試場");

    const MAX_SESSION_NUMBER = parseInt(configs["節數上限"]);
    const MAX_CLASSROOM_NUMBER = parseInt(configs["試場數量"]);
    const MAX_STUDENTS_PER_CLASSROOM = parseInt(configs["每間試場人數上限"]);
    const MAX_SUBJECTS_PER_CLASSROOM = parseInt(configs["試場容納科目上限"]);

    const sessions = getSessionStatistics();

    for (let i = 1; i < MAX_SESSION_NUMBER + 2; i++) {
        let students_sum = 0; // 用來加總同節次的所有試場人數
        const dgs = Object.entries(
            sessions[i].department_class_subject_statisics
        ).sort(descending_sorting);
        for (let j = 1; j < sessions[i].classrooms.length; j++) {
            let arranged_subjects = [];
            for (let k = 0; k < dgs.length; k++) {
                // 檢查此「班級-科目」是否已安排試場
                const has_duplicate = arranged_subjects.includes(dgs[k][0]);
                if (has_duplicate) {
                    continue;
                }

                // 檢查此試場是否還有名額
                const has_quota =
                    dgs[k][1] + sessions[i].classrooms[j].population <=
                    MAX_STUDENTS_PER_CLASSROOM;
                if (!has_quota) {
                    continue;
                }

                // 檢查此試場的科目數是否低於限制
                const under_subject_limitation =
                    1 +
                        Object.keys(
                            sessions[i].classrooms[j].class_subject_statisics
                        ).length <=
                    MAX_SUBJECTS_PER_CLASSROOM;
                if (!under_subject_limitation) {
                    continue;
                }

                if (
                    sessions[i].classrooms[j].population >=
                    MAX_STUDENTS_PER_CLASSROOM
                ) {
                    break;
                }

                if (!has_duplicate && has_quota && under_subject_limitation) {
                    sessions[i].students.forEach(function (row) {
                        let key = row[classNameColumn] + row[subject_column];
                        if (key == dgs[k][0] && row[classroom_column] == 0) {
                            row[classroom_column] = j;
                            sessions[i].classrooms[j].students.push(row);
                        }
                    });
                }

                arranged_subjects = arranged_subjects.concat(
                    Object.keys(
                        sessions[i].classrooms[j].class_subject_statisics
                    )
                );
            }
            students_sum += sessions[i].classrooms[j].population;
        }

        if (sessions[i].population != students_sum) {
            let merged_session_classrooms = [];
            for (let i = 1; i < MAX_SESSION_NUMBER + 2; i++) {
                for (let j = 1; j < sessions[i].classrooms.length; j++) {
                    merged_session_classrooms =
                        merged_session_classrooms.concat(
                            sessions[i].classrooms[j].students
                        );
                }
            }
            break;
        }
    }

    let modified_data = [];
    for (let i = 1; i < MAX_SESSION_NUMBER + 2; i++) {
        for (let j = 1; j < sessions[i].classrooms.length; j++) {
            modified_data = modified_data.concat(
                sessions[i].classrooms[j].students
            );
        }
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
            "(arrangeClassroom) 現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！"
        );
        SpreadsheetApp.getUi().alert(
            "現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！"
        );
    }

    if (sessions[9].students.length > 0) {
        SpreadsheetApp.getUi().alert(
            "部分考生被安排在第9節補考，請注意是否需要調整到中午應試！"
        );
    }

    filteredSheet.getRange("I:J").setNumberFormat("#,##0");
    sort_by_session_classroom();
}
