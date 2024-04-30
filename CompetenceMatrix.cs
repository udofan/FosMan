using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace FosMan {
    /// <summary>
    /// Описание матрицы компетенций
    /// </summary>
    public class CompetenceMatrix {
        /// <summary>
        /// Элементы матрицы
        /// </summary>
        public List<CompetenceMatrixItem> Items { get; set; }

        public bool IsLoaded { get => Items?.Any() ?? false; }
        /// <summary>
        /// Выявленные ошибки
        /// </summary>
        public List<string> Errors { get; set; }

        /// <summary>
        /// Загрузка матрицы из файла
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static CompetenceMatrix LoadFromFile(string fileName) {
            var matrix = new CompetenceMatrix() {
                Items = [],
                Errors = [],
            };

            try {
                using (var docx = DocX.Load(fileName)) {
                    if (docx.Tables.Count > 0) {
                        var table = docx.Tables[0];

                        TryParseTable(table, matrix);
                    }
                    else {
                        matrix.Errors.Add("В документе не найдено таблиц.");
                    }
                }
            }
            catch (Exception ex) {
                matrix.Errors.Add($"{ex.Message}\r\n{ex.StackTrace}");
            }

            return matrix;
        }

        /// <summary>
        /// Попытка отпарсить таблицу на таблицу компетенций
        /// </summary>
        /// <param name="table"></param>
        /// <param name="matrix"></param>
        /// <returns></returns>
        public static bool TryParseTable(Table table, CompetenceMatrix matrix) {
            CompetenceMatrixItem currItem = new();

            var matchedTable = false;   //флаг, что таблица с компетенциями

            for (var rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++) {
                var row = table.Rows[rowIdx];
                if (row.Cells.Count >= 3) {
                    //var cell0 = row.Cells[0];

                    var text = string.Join(" ", row.Cells[0].Paragraphs.Select(p => p.Text));
                    if (!string.IsNullOrEmpty(text)) { //проверка на очередной ряд с компетенцией
                        if (CompetenceMatrixItem.TryParse(text, out currItem)) {
                            matchedTable = true;
                            matrix.Items.Add(currItem);
                        }
                        else if (matchedTable) {
                            matrix.Errors.Add($"Не удалось распарсить текст [{text}] (ряд {rowIdx}, колонка 0).");
                        }
                    }

                    if (matchedTable) {
                        var achievement = new CompetenceAchievement();
                        currItem.Achievements.Add(achievement);

                        var text1 = string.Join(" ", row.Cells[1].Paragraphs.Select(p => p.Text));
                        var text2 = string.Join("\r\n", row.Cells[2].Paragraphs.Select(p => p.Text));
                        if (!achievement.TryParseIndicator(text1)) {
                            matrix.Errors.Add($"Не удалось распарсить текст [{text1}] (ряд {rowIdx}, колонка 1).");
                        }
                        if (!achievement.TryParseResult(text2)) {
                            matrix.Errors.Add($"Не удалось распарсить текст [{text2}] (ряд {rowIdx}, колонка 2).");
                        }
                    }
                }
                else {
                    matrix.Errors.Add("В таблице должно быть не менее 3 колонок.");
                    break;
                }
            }

            matrix.Check();

            return !matrix.Errors.Any();
        }

        /// <summary>
        /// Проверка матрицы
        /// </summary>
        private void Check() {
            foreach (var item in Items) {
                foreach (var achi in item.Achievements) {
                    if (string.IsNullOrEmpty(achi.Code)) {
                        Errors.Add($"Компетенция {item.Code}: индикатор не определён");
                    }
                    else if (!achi.Code.Contains(item.Code)) {
                        Errors.Add($"Компетенция {item.Code}: индикатор не соответствует компетенции - {achi.Code}");
                    }
                    if (string.IsNullOrEmpty(achi.ResultCode)) {
                        Errors.Add($"Компетенция {item.Code}: результат не определён");
                    }
                    else if (!achi.ResultCode.Contains(item.Code)) {
                        Errors.Add($"Компетенция {item.Code}: результат не соответствует компетенции - {achi.ResultCode}");
                    }
                }
            }
        }
    }
}
