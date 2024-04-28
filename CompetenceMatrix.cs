using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace FosMan {
    /// <summary>
    /// Описание матрицы компетенций
    /// </summary>
    public static class CompetenceMatrix {
        /// <summary>
        /// Элементы матрицы
        /// </summary>
        public static List<CompetenceMatrixItem> Items { get; set; }

        public static bool IsLoaded { get => Items?.Any() ?? false; }

        /// <summary>
        /// Загрузка матрицы из файла
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static void LoadFromFile(string fileName, out List<string> errors) {
            Items = [];
            errors = [];

            try {
                using (var docx = DocX.Load(fileName)) {
                    if (docx.Tables.Count > 0) {
                        var table = docx.Tables[0];

                        CompetenceMatrixItem currItem = new();

                        for (var rowIdx = 1; rowIdx < table.Rows.Count; rowIdx++) {
                            var row = table.Rows[rowIdx];
                            if (row.Cells.Count >= 3) {
                                var cell0 = row.Cells[0];

                                var text = cell0.Paragraphs[0].Text;
                                if (!string.IsNullOrEmpty(text)) { //ряд с очередной компетенцией
                                    currItem = new CompetenceMatrixItem() {
                                        Achievements = []
                                    };
                                    if (!currItem.InitFromText(text)) {
                                        errors.Add($"Не удалось распарсить текст [{text}].");
                                    }
                                    Items.Add(currItem);
                                }

                                var achievement = new CompetenceAchievement();
                                currItem.Achievements.Add(achievement);

                                var text1 = string.Join(" ", row.Cells[1].Paragraphs.Select(p => p.Text));
                                var text2 = string.Join("\r\n", row.Cells[2].Paragraphs.Select(p => p.Text));
                                achievement.SetIndicator(text1);
                                achievement.SetResult(text2);
                            }
                            else {
                                errors.Add("В таблице должно быть не менее 3 колонок.");
                                break;
                            }
                        }
                    }
                    else {
                        errors.Add("В документе не найдено таблиц.");
                    }
                }
            }
            catch (Exception ex) {
                errors.Add(ex.Message);
            }
            //if (errors.Any()) {
            //    matrix = null;
            //}

            return;
        }
    }
}
