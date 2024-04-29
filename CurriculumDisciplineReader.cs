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
        static List<CurriculumDisciplineHeader> m_headers { get; } = [
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
                Text = "КОД",
                TestFunction = EPropertyTestFunction.Equals,
                TargetProperty = "DepartmentCode",
                TargetType = typeof(string)
            },
            new CurriculumDisciplineHeader() {
                Text = "НАИМЕНОВАНИЕ",
                TestFunction = EPropertyTestFunction.Contains,
                TargetProperty = "Department",
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
                        foreach (var header in m_headers) {
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
                    foreach (var header in m_headers) {
                        if (!detectedColumns.ContainsValue(header)) {
                            curriculum.Errors.Add($"Не удалось найти заголовок [{header.Text}] для свойства {header.TargetProperty}");
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
                        discipline.SetProperty(item.Value.TargetProperty, item.Value.TargetType, cellValue);
                    }
                }
                if (discipline != null) {
                    if (curriculum.Disciplines.TryAdd(discipline.Key, discipline)) {
                        if (discipline.Type == EDisciplineType.Unknown) {
                            curriculum.Errors.Add($"Не удалось определить тип дисциплины [{discipline.Name}] по индексу [{discipline.Index}] (строка {rowIdx})");
                        }
                    }
                    else {
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
