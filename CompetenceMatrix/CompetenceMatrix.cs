using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Xceed.Document.NET;
using Xceed.Words.NET;
using static FosMan.Enums;

namespace FosMan {
    /// <summary>
    /// Описание матрицы компетенций
    /// </summary>
    public class CompetenceMatrix {
        //вариант РПД
        //Код и наименование компетенций - Коды и индикаторы достижения компетенций - Коды и результаты обучения
        static Regex m_regexCompetenceRpdHeader0 = new(@"код.*наименование", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        static Regex m_regexCompetenceRpdHeader1 = new(@"код.*индикатор", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        static Regex m_regexCompetenceRpdHeader2 = new(@"код.*результат", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        //вариант ФОС - таблица 2.1
        //Код и наименование компетенций - Код и индикаторы достижения компетенций - Этапы формирования компетенций (семестр)
        static Regex m_regexCompetenceFos21Header0 = new(@"код.*наименование", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        static Regex m_regexCompetenceFos21Header1 = new(@"код.*индикатор", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        static Regex m_regexCompetenceFos21Header2 = new(@"этап", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        //вариант ФОС - таблица 2.2
        //Код компетенции - Коды индикаторов достижения компетенций - Коды и результаты обучения
        static Regex m_regexCompetenceFos22Header0 = new(@"код.*компетенц", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        static Regex m_regexCompetenceFos22Header1 = new(@"код.*индикатор", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        static Regex m_regexCompetenceFos22Header2 = new(@"код.*результат", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Элементы матрицы
        /// </summary>
        [JsonInclude]
        public List<CompetenceMatrixItem> Items { get; set; }

        [JsonIgnore]
        public bool IsLoaded { get => Items?.Any() ?? false; }
        /// <summary>
        /// Выявленные ошибки
        /// </summary>
        [JsonIgnore]
        public ErrorList Errors { get; set; }

        /// <summary>
        /// Флаг, что матрица загружена полностью
        /// </summary>
        [JsonIgnore]
        public bool IsComplete { get => Items?.FirstOrDefault()?.Achievements?.FirstOrDefault()?.Results?.FirstOrDefault() != null; }

        /// <summary>
        /// Загрузка матрицы из файла
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static CompetenceMatrix LoadFromFile(string fileName) {
            var matrix = new CompetenceMatrix() {
                Items = [],
                Errors = new(),
            };

            try {
                using (var docx = DocX.Load(fileName)) {
                    if (docx.Tables.Count > 0) {
                        foreach (var table in docx.Tables) {
                            //var table = docx.Tables[1];
                            TryParseTable(table, ECompetenceMatrixFormat.Rpd, matrix);
                            //break;
                        }
                        matrix?.Check();
                    }
                    else {
                        matrix.Errors.Add(EErrorType.CompetenceMatrixNoTables);
                    }
                }
            }
            catch (Exception ex) {
                matrix.Errors.AddException(ex); // ($"{ex.Message}\r\n{ex.StackTrace}");
            }

            return matrix;
        }

        /// <summary>
        /// Поиск элемента по коду компетенции
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        CompetenceMatrixItem FindItem(string code) {
            var normalizedCode = CompetenceMatrixItem.NormalizeCode(code);

            return Items.FirstOrDefault(i => i.Code.Equals(normalizedCode));
        }

        /// <summary>
        /// Попытка отпарсить таблицу на таблицу компетенций
        /// </summary>
        /// <param name="table"></param>
        /// <param name="matrix"></param>
        /// <returns></returns>
        public static bool TryParseTable(Table table, ECompetenceMatrixFormat format, CompetenceMatrix matrix) {
            CompetenceMatrixItem currItem = null;
            CompetenceAchievement currAchievement = null;

            var matchedTable = false;   //флаг, что таблица с компетенциями

            for (var rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++) {
                var row = table.Rows[rowIdx];
                if (row.Cells.Count >= 3) {
                    //ячейка 0
                    var textCompetence = string.Join(" ", row.Cells[0].Paragraphs.Select(p => p.Text));
                    if (!string.IsNullOrEmpty(textCompetence) && CompetenceMatrixItem.TestCode(textCompetence)) { //проверка на очередной ряд с компетенцией
                        if (format == ECompetenceMatrixFormat.Fos22) {
                            //к этому моменту в матрице уже должны быть ряды и индикаторы достижений
                            //ищем currItem по тексту textCompetence
                            currItem = matrix.FindItem(textCompetence);
                            if (currItem != null) {
                                matchedTable = true;
                            }
                            else if (currItem == null) {
                                matrix.Errors.Add(EErrorType.CompetenceMatrixMissingItem, "код [{textCompetence}] (ряд {rowIdx}, колонка 0)");
                            }
                        }
                        else { //в случае других форматов таблицы
                            if (CompetenceMatrixItem.TryParse(textCompetence, out currItem)) {
                                matchedTable = true;
                                matrix.Items.Add(currItem);
                            }
                            else if (matchedTable) {
                                matrix.Errors.Add(EErrorType.CompetenceMatrixItemParseError, $"значение [{textCompetence}] (ряд {rowIdx}, колонка 0)");
                            }
                        }
                    }

                    if (matchedTable) {
                        //ячейка 1
                        var textIndicator = string.Join(" ", row.Cells[1].Paragraphs.Select(p => p.Text)).Trim();
                        if (!string.IsNullOrEmpty(textIndicator)) { //значение индикатора есть?
                            if (format == ECompetenceMatrixFormat.Fos22) {
                                //к этому моменту в матрице уже должны быть ряды и индикаторы достижений
                                //ищем currAchievement по тексту textIndicator
                                currAchievement = currItem.FindAchievement(textIndicator);
                                if (currAchievement == null) {
                                    matrix.Errors.Add(EErrorType.CompetenceMatrixMissingAchievement, $"код [{textIndicator}] (ряд {rowIdx}, колонка 1)");
                                }
                            }
                            else {  //в случае других форматов таблицы
                                if (CompetenceAchievement.TryParseIndicator(textIndicator, out currAchievement)) {
                                    currItem.Achievements.Add(currAchievement);
                                }
                                else {
                                    matrix.Errors.Add(EErrorType.CompetenceMatrixAchievementParseError, $"значение [{textIndicator}] (ряд {rowIdx}, колонка 1)");
                                }
                            }
                        }
                        //ячейка 2
                        if (format == ECompetenceMatrixFormat.Rpd || format == ECompetenceMatrixFormat.Fos22) {
                            var textResult = string.Join("\r\n", row.Cells[2].Paragraphs.Select(p => p.Text)).Trim();
                            if (!string.IsNullOrEmpty(textResult)) {
                                if (CompetenceResult.TryParse(textResult, out var result)) {
                                    currAchievement?.Results.Add(result);
                                }
                                else {
                                    matrix.Errors.Add(EErrorType.CompetenceMatrixResultParseError, $"значение [{textResult}] (ряд {rowIdx}, колонка 2)");
                                }
                            }
                            else {
                                matrix.Errors.Add(EErrorType.CompetenceMatrixMissingResult, $"(ряд {rowIdx}, колонка 2)");
                            }
                        }
                        else if (format == ECompetenceMatrixFormat.Fos21) {
                            //в случае формата ФОС - 2.1 во 2-ой ячейке лежит номер семестра
                            //if (row.Cells[2].RowSpan == 0) {
                            var stage = row.Cells[2].GetText();
                            if (!string.IsNullOrEmpty(stage)) {
                                if (int.TryParse(stage, out var intValue)) {
                                    currItem.Semester = intValue;
                                }
                                else {
                                    matrix.Errors.Add(EErrorType.CompetenceMatrixSemesterParseError, $"значение [{stage}] (ряд {rowIdx}, колонка 2)");
                                }
                            }
                            else if (currItem.Semester < 0) {
                                matrix.Errors.Add(EErrorType.CompetenceMatrixMissingSemester, $"(ряд {rowIdx}, колонка 2)");
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
                Errors.Add(EErrorType.CompetenceMatrixIsEmpty);
            }

            foreach (var item in Items) {
                if (!item.Achievements.Any()) {
                    Errors.Add(EErrorType.CompetenceMatrixItemMissingAchievements, $"Компетенция {item.Code}");
                }
                foreach (var achi in item.Achievements) {
                    if (string.IsNullOrEmpty(achi.Code)) {
                        Errors.Add(EErrorType.CompetenceMatrixItemMissingAchievementCode, $"Компетенция {item.Code}");
                    }
                    else if (!achi.Code.Contains(item.Code)) {
                        Errors.Add(EErrorType.CompetenceMatrixAchievementCodeMismatch, $"Компетенция {item.Code} - индикатор {achi.Code}");
                    }
                    foreach (var res in achi.Results) {
                        if (string.IsNullOrEmpty(res.Code)) {
                            Errors.Add(EErrorType.CompetenceMatrixMissingAchievementResult, $"Компетенция {item.Code} - индикатор {achi.Code}");
                        }
                        else {
                            if (!res.Code.Contains(item.Code)) {
                                Errors.Add(EErrorType.CompetenceMatrixResultCodeItemMismatch, $"Компетенция {item.Code} - результат индикатора {res.Code}");
                            }
                            if (!res.Code.Contains(achi.Code)) {
                                Errors.Add(EErrorType.CompetenceMatrixResultCodeAchievementMismatch, $"Компетенция {item.Code} - индикатор {achi.Code} - результат индикатора {res.Code}");
                            }
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
            if ((Errors?.IsEmpty ?? true) == false) {
                html += "<div style='color: red;'><b>ОШИБКИ:</b></div>";
                html += string.Join("", Errors.Items.Select(e => $"<div style='color: red'>{e}</div>"));
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
        /// Список кодов всех результатов достижений матрицы
        /// </summary>
        /// <returns></returns>
        public HashSet<string> GetAllResultCodes(List<CompetenceMatrixItem> items = null) {
            var achiList = new List<CompetenceAchievement>();
            items ??= Items;
            items.ForEach(x => achiList.AddRange(x.Achievements));
            var resultCodeList = new List<string>();
            achiList.ForEach(x => resultCodeList.AddRange(x.Results.Select(a => a.Code)));
            return [.. resultCodeList];
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
        static internal bool TestTable(Table table, out ECompetenceMatrixFormat format) {
            format = ECompetenceMatrixFormat.Unknown;

            if (table.RowCount > 0 && table.ColumnCount >= 3) {
                var row = table.Rows[0];
                var header0 = row.Cells[0].GetText();
                var header1 = row.Cells[1].GetText();
                var header2 = row.Cells[2].GetText();
                //вариант РПД
                if (m_regexCompetenceRpdHeader0.IsMatch(header0) &&
                    m_regexCompetenceRpdHeader1.IsMatch(header1) &&
                    m_regexCompetenceRpdHeader2.IsMatch(header2)) {
                    format = ECompetenceMatrixFormat.Rpd;
                }
                //вариант ФОС-2.1
                else if (m_regexCompetenceFos21Header0.IsMatch(header0) &&
                         m_regexCompetenceFos21Header1.IsMatch(header1) &&
                         m_regexCompetenceFos21Header2.IsMatch(header2)) {
                    format = ECompetenceMatrixFormat.Fos21;
                }
                //вариант ФОС-2.2
                else if (m_regexCompetenceFos22Header0.IsMatch(header0) &&
                         m_regexCompetenceFos22Header1.IsMatch(header1) &&
                         m_regexCompetenceFos22Header2.IsMatch(header2)) {
                    format = ECompetenceMatrixFormat.Fos22;
                }
            }

            return format != ECompetenceMatrixFormat.Unknown;
        }

        /// <summary>
        /// Получить список достижений по списку результатов
        /// </summary>
        /// <param name="resultCodes"></param>
        /// <returns></returns>
        public HashSet<CompetenceAchievement> GetAchievements(HashSet<string> resultCodes) {
            HashSet<CompetenceAchievement> list = new();

            foreach (var item in Items) {
                foreach (var achi in item.Achievements) {
                    foreach (var res in achi.Results) {
                        if (resultCodes.Contains(res.Code)) {
                            list.Add(achi);
                            break;
                        }
                    }
                }
            }

            return list;
        }
    }
}
