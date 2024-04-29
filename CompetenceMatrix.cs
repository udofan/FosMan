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
        public void LoadFromFile(string fileName) {
            Items = [];
            Errors = [];

            try {
                using (var docx = DocX.Load(fileName)) {
                    if (docx.Tables.Count > 0) {
                        var table = docx.Tables[0];

                        CompetenceMatrixItem currItem = new();

                        for (var rowIdx = 1; rowIdx < table.Rows.Count; rowIdx++) {
                            var row = table.Rows[rowIdx];
                            if (row.Cells.Count >= 3) {
                                //var cell0 = row.Cells[0];

                                var text = string.Join(" ", row.Cells[0].Paragraphs.Select(p => p.Text));
                                if (!string.IsNullOrEmpty(text)) { //ряд с очередной компетенцией
                                    currItem = new CompetenceMatrixItem() {
                                        Achievements = []
                                    };
                                    if (!currItem.InitFromText(text)) {
                                        Errors.Add($"Не удалось распарсить текст [{text}].");
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
                                Errors.Add("В таблице должно быть не менее 3 колонок.");
                                break;
                            }
                        }
                    }
                    else {
                        Errors.Add("В документе не найдено таблиц.");
                    }
                }
            }
            catch (Exception ex) {
                Errors.Add($"{ex.Message}\r\n{ex.StackTrace}");
            }

            return;
        }
    }
}
