using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace FosMan {
    /// <summary>
    /// Описание матрицы компетенций
    /// </summary>
    public class CompetenceMatrix {
        static Regex m_regexCompetenceHeader0 = new(@"код.*наим.*компетенц", RegexOptions.IgnoreCase | RegexOptions.Compiled);

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
                        foreach (var table in docx.Tables) {
                            //var table = docx.Tables[1];
                            TryParseTable(table, matrix);
                            //break;
                        }
                        matrix?.Check();
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
            CompetenceMatrixItem currItem = null;
            CompetenceAchievement currAchievement = null;

            var matchedTable = false;   //флаг, что таблица с компетенциями

            for (var rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++) {
                var row = table.Rows[rowIdx];
                if (row.Cells.Count >= 3) {
                    var textCompetence = string.Join(" ", row.Cells[0].Paragraphs.Select(p => p.Text));
                    if (!string.IsNullOrEmpty(textCompetence)) { //проверка на очередной ряд с компетенцией
                        if (CompetenceMatrixItem.TryParse(textCompetence, out currItem)) {
                            matchedTable = true;
                            matrix.Items.Add(currItem);
                        }
                        else if (matchedTable) {
                            matrix.Errors.Add($"Не удалось распарсить текст компетенции [{textCompetence}] (ряд {rowIdx}, колонка 0).");
                        }
                    }

                    if (matchedTable) {
                        var textIndicator = string.Join(" ", row.Cells[1].Paragraphs.Select(p => p.Text)).Trim();
                        if (!string.IsNullOrEmpty(textIndicator)) { //значение индикатора есть?
                            if (CompetenceAchievement.TryParseIndicator(textIndicator, out currAchievement)) {
                                currItem.Achievements.Add(currAchievement);
                            }
                            else {
                                matrix.Errors.Add($"Не удалось распарсить текст индикатора [{textIndicator}] (ряд {rowIdx}, колонка 1).");
                            }

                        }

                        var textResult = string.Join("\r\n", row.Cells[2].Paragraphs.Select(p => p.Text)).Trim();
                        if (!string.IsNullOrEmpty(textResult)) {
                            if (CompetenceResult.TryParse(textResult, out var result)) {
                                currAchievement?.Results.Add(result);
                            }
                            else {
                                matrix.Errors.Add($"Не удалось распарсить текст результата [{textResult}] (ряд {rowIdx}, колонка 2).");
                            }
                        }
                    }
                }
                else {
                    //matrix.Errors.Add("В таблице компетенций должно быть не менее 3 колонок.");
                    break;
                }
            }

            return matrix.Items.Any();
        }

        /// <summary>
        /// Проверка матрицы
        /// </summary>
        public void Check() {
            if (!Items.Any()) {
                Errors.Add("Список компетенций не определён");
            }

            foreach (var item in Items) {
                if (!item.Achievements.Any()) {
                    Errors.Add($"Компетенция {item.Code}: список индикаторов не определён");
                }
                foreach (var achi in item.Achievements) {
                    if (string.IsNullOrEmpty(achi.Code)) {
                        Errors.Add($"Компетенция {item.Code}: индикатор не определён");
                    }
                    else if (!achi.Code.Contains(item.Code)) {
                        Errors.Add($"Компетенция {item.Code}: индикатор не соответствует компетенции - {achi.Code}");
                    }
                    foreach (var res in achi.Results) {
                        if (string.IsNullOrEmpty(res.Code)) {
                            Errors.Add($"Компетенция {item.Code}: результат не определён для индикатора {achi.Code}");
                        }
                        else if (!res.Code.Contains(item.Code)) {
                            Errors.Add($"Компетенция {item.Code}: результат индикатора {achi.Code} не соответствует компетенции - {res.Code}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Создать html-отчет
        /// </summary>
        /// <returns></returns>
        public string CreateHtmlReport() {
            var html = "<html><body>";
            var tdStyle = " style='border: 1px solid; vertical-align: top;'";

            //если есть ошибки
            if (Errors?.Any() ?? false) {
                html += "<div style='color: red;'><b>ОШИБКИ:</b></div>";
                html += string.Join("", Errors.Select(e => $"<div style='color: red'>{e}</div>"));
            }

            //формирование матрицы
            html += "<table style='border: 1px solid'>";
            html += $"<tr><th {tdStyle}><b>Код и наименование компетенций</b></th>" +
                        $"<th {tdStyle}><b>Коды и индикаторы достижения компетенций</b></th>" +
                        $"<th {tdStyle}><b>Коды и результаты обучения</b><</th></tr>";
            foreach (var item in Items) {
                //var item = Items[3]; //отладка
                html += "<tr>";
                var rowSpan = item.Achievements.Sum(a => Math.Max(a.Results.Count, 1));
                html += $"<td {tdStyle} {(rowSpan > 1 ? $"rowspan='{rowSpan}'" : "")}><b>{item.Code}</b>. {item.Title}</td>";

                for (var achiIdx = 0; achiIdx < item.Achievements.Count; achiIdx++) {
                    var achi = item.Achievements[achiIdx];
                    if (achiIdx > 0) html += $"<tr>";
                    var achiCellRowSpan = achi.Results.Count > 1 ? $"rowspan='{achi.Results.Count}'" : "";
                    html += $"<td {tdStyle} {achiCellRowSpan}'>{achi.Code}. {achi.Indicator}</td>";

                    for (var resIdx = 0; resIdx < achi.Results.Count; resIdx++) {
                        var res = achi.Results[resIdx];
                        if (resIdx > 0) html += "<tr>";
                        html += $"<td {tdStyle}>{res.Code}:<br />{res.Description}</td>";
                        //if (resIdx > 0) 
                        html += "</tr>";
                    }
                    if (achi.Results.Count == 0) {
                        html += $"<td {tdStyle}>???</td>";
                        html += "</tr>";
                    }
                    //if (achiIdx > 0) 
                    //html += "</tr>";
                }
                //html += "</tr>";
                //break; //отладка
            }
            html += "</table>";

            html += "</body></html>";

            return html;
        }

        /// <summary>
        /// Список кодов всех достижений матрицы
        /// </summary>
        /// <returns></returns>
        public HashSet<string> GetAllAchievementCodes(List<CompetenceMatrixItem> items = null) {
            var achiCodeList = new List<string>();
            items ??= Items;
            items.ForEach(x => achiCodeList.AddRange(x.Achievements.Select(a => a.Code)));
            return [.. achiCodeList];
        }

        /// <summary>
        /// Получить список элементов матрицы по списку кодов индикаторов
        /// </summary>
        /// <param name="achievementCodes"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        internal List<CompetenceMatrixItem> GetItems(HashSet<string> achievementCodes) {
            return Items?.Where(i => i.Achievements?.Any(a => achievementCodes?.Contains(a.Code) ?? false) ?? false).ToList();
        }

        /// <summary>
        /// Проверка, что передана таблица компетенций
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        internal bool TestTable(Table table) {
            var result = false;

            if (table.RowCount > 0 && table.ColumnCount >= 3) {
                var row = table.Rows[0];
                var header0 = row.Cells[0].GetText();
                if (m_regexCompetenceHeader0.IsMatch(header0)) {
                    result = true;
                }
            }

            return result;
        }
    }
}
