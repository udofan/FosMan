using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    static internal class CurriculumDisciplineReader {
        /// <summary>
        /// Описание искомых колонок в Excel-листе
        /// </summary>
        public static List<CurriculumDisciplineHeader> DefaultHeaders { get; } = [
            new CurriculumDisciplineHeader() {
                Text = "ИНДЕКС",
                TestFunction = EPropertyTestFunction.Contains,
                TargetProperty = "Index",
                TargetType = typeof(string)
            },
            new CurriculumDisciplineHeader() {
                Text = "НАИМЕНОВАНИЕ",
                TestFunction = EPropertyTestFunction.Contains,
                TargetProperty = "Name",
                TargetType = typeof(string)
            },
            new CurriculumDisciplineHeader() {
                Text = "ЭКЗА",
                TestFunction = EPropertyTestFunction.Contains,
                TargetProperty = "ControlFormExamHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "ЗАЧЕТ",
                TestFunction = EPropertyTestFunction.Contains,
                TargetProperty = "ControlFormTestHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "ЗАЧЕТ С ОЦ",
                TestFunction = EPropertyTestFunction.Contains,
                TargetProperty = "ControlFormTestWithAGradeHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "КР",
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "ControlFormControlWorkHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "По плану", 
                TestFunction = EPropertyTestFunction.Contains, 
                TargetProperty = "TotalByPlanHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт. раб.",
                TestFunction = EPropertyTestFunction.Contains,
                TargetProperty = "TotalContactWorkHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "СР",
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "TotalSelfStudyHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт",
                TestFunction = EPropertyTestFunction.StartsWith,
                TargetProperty = "TotalControlHours",
                TargetType = typeof(int)
            },
            //Семестр 1
            new CurriculumDisciplineHeader() {
                Text = "Итого",
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                TargetType = typeof(int), 
                SubProperty = "TotalHours",
                PropertyIndex = 0
            },
            new CurriculumDisciplineHeader() {
                Text = "Лек",
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 0,
                SubProperty = "LectureHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Пр",
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 0,
                SubProperty = "PracticalHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Ср",
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 0,
                SubProperty = "SelfStudyHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт",
                MatchNumber = 1,
                TestFunction = EPropertyTestFunction.StartsWith,
                TargetProperty = "Semesters",
                PropertyIndex = 0,
                SubProperty = "ControlHours",
                TargetType = typeof(int)
            },
            //Семестр 2
            new CurriculumDisciplineHeader() {
                Text = "Итого",
                MatchNumber = 1,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                TargetType = typeof(int),
                SubProperty = "TotalHours",
                PropertyIndex = 1
            },
            new CurriculumDisciplineHeader() {
                Text = "Лек",
                MatchNumber = 1,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 1,
                SubProperty = "LectureHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Пр",
                MatchNumber = 1,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 1,
                SubProperty = "PracticalHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Ср",
                MatchNumber = 1,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 1,
                SubProperty = "SelfStudyHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт",
                MatchNumber = 2,
                TestFunction = EPropertyTestFunction.StartsWith,
                TargetProperty = "Semesters",
                PropertyIndex = 1,
                SubProperty = "ControlHours",
                TargetType = typeof(int)
            },
            //Семестр 3
            new CurriculumDisciplineHeader() {
                Text = "Итого",
                MatchNumber = 2,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                TargetType = typeof(int),
                SubProperty = "TotalHours",
                PropertyIndex = 2
            },
            new CurriculumDisciplineHeader() {
                Text = "Лек",
                MatchNumber = 2,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 2,
                SubProperty = "LectureHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Пр",
                MatchNumber = 2,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 2,
                SubProperty = "PracticalHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Ср",
                MatchNumber = 2,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 2,
                SubProperty = "SelfStudyHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт",
                MatchNumber = 3,
                TestFunction = EPropertyTestFunction.StartsWith,
                TargetProperty = "Semesters",
                PropertyIndex = 2,
                SubProperty = "ControlHours",
                TargetType = typeof(int)
            },
            //Семестр 4
            new CurriculumDisciplineHeader() {
                Text = "Итого",
                MatchNumber = 3,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                TargetType = typeof(int),
                SubProperty = "TotalHours",
                PropertyIndex = 3
            },
            new CurriculumDisciplineHeader() {
                Text = "Лек",
                MatchNumber = 3,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 3,
                SubProperty = "LectureHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Пр",
                MatchNumber = 3,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 3,
                SubProperty = "PracticalHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Ср",
                MatchNumber = 3,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 3,
                SubProperty = "SelfStudyHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт",
                MatchNumber = 4,
                TestFunction = EPropertyTestFunction.StartsWith,
                TargetProperty = "Semesters",
                PropertyIndex = 3,
                SubProperty = "ControlHours",
                TargetType = typeof(int)
            },
            //Семестр 5
            new CurriculumDisciplineHeader() {
                Text = "Итого",
                MatchNumber = 4,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                TargetType = typeof(int),
                SubProperty = "TotalHours",
                PropertyIndex = 4
            },
            new CurriculumDisciplineHeader() {
                Text = "Лек",
                MatchNumber = 4,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 4,
                SubProperty = "LectureHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Пр",
                MatchNumber = 4,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 4,
                SubProperty = "PracticalHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Ср",
                MatchNumber = 4,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 4,
                SubProperty = "SelfStudyHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт",
                MatchNumber = 5,
                TestFunction = EPropertyTestFunction.StartsWith,
                TargetProperty = "Semesters",
                PropertyIndex = 4,
                SubProperty = "ControlHours",
                TargetType = typeof(int)
            },
            //Семестр 6
            new CurriculumDisciplineHeader() {
                Text = "Итого",
                MatchNumber = 5,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                TargetType = typeof(int),
                SubProperty = "TotalHours",
                PropertyIndex = 5
            },
            new CurriculumDisciplineHeader() {
                Text = "Лек",
                MatchNumber = 5,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 5,
                SubProperty = "LectureHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Пр",
                MatchNumber = 5,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 5,
                SubProperty = "PracticalHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Ср",
                MatchNumber = 5,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 5,
                SubProperty = "SelfStudyHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт",
                MatchNumber = 6,
                TestFunction = EPropertyTestFunction.StartsWith,
                TargetProperty = "Semesters",
                PropertyIndex = 5,
                SubProperty = "ControlHours",
                TargetType = typeof(int)
            },
            //Семестр 7
            new CurriculumDisciplineHeader() {
                Text = "Итого",
                MatchNumber = 6,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                TargetType = typeof(int),
                SubProperty = "TotalHours",
                PropertyIndex = 6
            },
            new CurriculumDisciplineHeader() {
                Text = "Лек",
                MatchNumber = 6,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 6,
                SubProperty = "LectureHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Пр",
                MatchNumber = 6,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 6,
                SubProperty = "PracticalHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Ср",
                MatchNumber = 6,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 6,
                SubProperty = "SelfStudyHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт",
                MatchNumber = 7,
                TestFunction = EPropertyTestFunction.StartsWith,
                TargetProperty = "Semesters",
                PropertyIndex = 6,
                SubProperty = "ControlHours",
                TargetType = typeof(int)
            },
            //Семестр 8
            new CurriculumDisciplineHeader() {
                Text = "Итого",
                MatchNumber = 7,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                TargetType = typeof(int),
                SubProperty = "TotalHours",
                PropertyIndex = 7
            },
            new CurriculumDisciplineHeader() {
                Text = "Лек",
                MatchNumber = 7,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 7,
                SubProperty = "LectureHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Пр",
                MatchNumber = 7,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 7,
                SubProperty = "PracticalHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Ср",
                MatchNumber = 7,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 7,
                SubProperty = "SelfStudyHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт",
                MatchNumber = 8,
                TestFunction = EPropertyTestFunction.StartsWith,
                TargetProperty = "Semesters",
                PropertyIndex = 7,
                SubProperty = "ControlHours",
                TargetType = typeof(int)
            },
            //Семестр 9
            new CurriculumDisciplineHeader() {
                Text = "Итого",
                MatchNumber = 8,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                TargetType = typeof(int),
                SubProperty = "TotalHours",
                PropertyIndex = 8
            },
            new CurriculumDisciplineHeader() {
                Text = "Лек",
                MatchNumber = 8,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 8,
                SubProperty = "LectureHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Пр",
                MatchNumber = 8,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 8,
                SubProperty = "PracticalHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Ср",
                MatchNumber = 8,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 8,
                SubProperty = "SelfStudyHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт",
                MatchNumber = 9,
                TestFunction = EPropertyTestFunction.StartsWith,
                TargetProperty = "Semesters",
                PropertyIndex = 8,
                SubProperty = "ControlHours",
                TargetType = typeof(int)
            },
            //Семестр 10
            new CurriculumDisciplineHeader() {
                Text = "Итого",
                MatchNumber = 9,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                TargetType = typeof(int),
                SubProperty = "TotalHours",
                PropertyIndex = 9
            },
            new CurriculumDisciplineHeader() {
                Text = "Лек",
                MatchNumber = 9,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 9,
                SubProperty = "LectureHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Пр",
                MatchNumber = 9,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 9,
                SubProperty = "PracticalHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Ср",
                MatchNumber = 9,
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "Semesters",
                PropertyIndex = 9,
                SubProperty = "SelfStudyHours",
                TargetType = typeof(int)
            },
            new CurriculumDisciplineHeader() {
                Text = "Конт",
                MatchNumber = 10,
                TestFunction = EPropertyTestFunction.StartsWith,
                TargetProperty = "Semesters",
                PropertyIndex = 9,
                SubProperty = "ControlHours",
                TargetType = typeof(int)
            },

            new CurriculumDisciplineHeader() {
                Text = "КОД",
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "DepartmentCode",
                TargetType = typeof(string)
            },
            new CurriculumDisciplineHeader() {
                Text = "НАИМЕНОВАНИЕ",
                TestFunction = EPropertyTestFunction.Contains,
                TargetProperty = "DepartmentName",
                TargetType = typeof(string),
                MatchNumber = 1
            },
            new CurriculumDisciplineHeader() {
                Text = "КОМПЕТЕНЦ",
                TestFunction = EPropertyTestFunction.Contains,
                TargetProperty = "Competences",
                TargetType = typeof(string)
            }
        ];

        /// <summary>
        /// Проверка ряда таблицы на заголовки
        /// </summary>
        /// <param name="row"></param>
        /// <param name="detectedColumns"></param>
        /// <returns></returns>
        public static bool TestHeaderRow(DataTable table, int rowIdx,
                                         Curriculum curriculum,
                                         out Dictionary<int, CurriculumDisciplineHeader> detectedColumns) {
            var result = false;
            
            detectedColumns = null;

            try {
                var row = table.Rows[rowIdx];
                detectedColumns = [];

                for (var colIdx = 0; colIdx < table.Columns.Count; colIdx++) {
                    var cellValue = row[colIdx] as string;

                    if (!string.IsNullOrEmpty(cellValue)) {
                        foreach (var header in App.Config.CurriculumDisciplineParseItems) { // m_headers) {
                            if (header.Match(cellValue)) {
                                //кол-во заголовков с одинаковым текстом
                                var matchCount = detectedColumns.Values.Where(x => x.Text == header.Text).Count();
                                if (header.MatchNumber == matchCount) {
                                    detectedColumns[colIdx] = header;
                                    break;
                                }
                            }
                        }
                    }
                }

                result = detectedColumns.Count > 0;

                if (result) {
                    //проверка, что какие-то заголовки не нашли в таблице
                    foreach (var header in App.Config.CurriculumDisciplineParseItems) { // m_headers) {
                        if (!detectedColumns.ContainsValue(header)) {
                            var err = $"Не удалось найти заголовок [{header.Text}] для свойства {header.TargetProperty}";
                            if (!string.IsNullOrEmpty(header.SubProperty) && header.PropertyIndex >= 0) {
                                err += $".{header.SubProperty}[{header.PropertyIndex}]";
                            }
                            curriculum.Errors.Add(err);
                        }
                    }
                }
            }
            catch (Exception ex) {
                curriculum.Errors.Add($"{ex.Message}\r\n{ex.StackTrace}");
                result = false;
            }

            return result;
        }

        /// <summary>
        /// Обработка значащего ряда
        /// </summary>
        /// <param name="tablePlan"></param>
        /// <param name="rowIdx"></param>
        /// <param name="curriculum"></param>
        /// <exception cref="NotImplementedException"></exception>
        internal static void ProcessRow(DataTable tablePlan, int rowIdx, Dictionary<int, CurriculumDisciplineHeader> headers, Curriculum curriculum) {
            try {
                var row = tablePlan.Rows[rowIdx];

                CurriculumDiscipline discipline = null;
                foreach (var item in headers) {
                    var cellValue = row[item.Key] as string;

                    if (!string.IsNullOrEmpty(cellValue)) {
                        discipline ??= new();
                        discipline.SetProperty(item.Value.TargetProperty, item.Value.TargetType, cellValue, item.Value.PropertyIndex, item.Value.SubProperty);
                    }
                }
                if (discipline?.Check(curriculum, rowIdx) ?? false) {
                    if (!curriculum.AddDiscipline(discipline)) {
                        curriculum.Errors.Add($"Обнаружено повторное упоминание дисциплины [{discipline.Name}] на строке {rowIdx}");
                    }
                }
            }
            catch (Exception ex) {
                curriculum.Errors.Add($"{ex.Message}\r\n{ex.StackTrace}");
            }
        }
    }
}
