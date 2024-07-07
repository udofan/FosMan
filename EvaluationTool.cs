using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;
using static FosMan.Enums;
using System.Xml.Linq;

namespace FosMan {
    /// <summary>
    /// Описание оценочного средства
    /// </summary>
    public class EvaluationTool : BaseObj {
        //выражение для отлова вопросов в тестах
        static Regex m_testingQuestionMark = new(@"^(\d*[\.]*\s*вопрос|\d+[\.]*\s+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Тип
        /// </summary>
        [JsonInclude]
        public EEvaluationTool Type { get; set; }
        /// <summary>
        /// Номер раздела в ФОСе
        /// </summary>
        [JsonInclude]
        public string ChapterNum { get; set; }
        /// <summary>
        /// Индекс таблицы
        /// </summary>
        [JsonInclude]
        public int TableIndex { get; set; }
        /// <summary>
        /// Объект таблицы
        /// </summary>
        [JsonIgnore]
        public Table Table { get; set; }
        /// <summary>
        /// Номер колонки со списком элементов
        /// </summary>
        [JsonIgnore]
        public int TableColIndexItems { get; set; }
        /// <summary>
        /// Номер колонки со списком индикаторов компетенций
        /// </summary>
        [JsonIgnore]
        public int TableColIndexCompetenceIndicators { get; set; }
        /// <summary>
        /// Прямой список
        /// </summary>
        [JsonInclude]
        public List<string> ListItems { get; set; }
        /// <summary>
        /// Элементы теста
        /// </summary>
        [JsonInclude]
        public List<TestingItem> TestingItems { get; set; }
        [JsonInclude]
        public List<string> XmlTestingItems { get; set; }
        /// <summary>
        /// Список индикаторов компетенций
        /// </summary>
        [JsonInclude]
        public HashSet<string> CompetenceIndicators { get; set; }

        /// <summary>
        /// Парсинг элементов оценочного средства
        /// </summary>
        public void ParseItems(Fos fos) {
            if (Table != null) {
                for (var r = 1; r < Table.RowCount; r++) {
                    var row = Table.Rows[r];
                    if (row.Cells.Count > TableColIndexCompetenceIndicators) {
                        CompetenceIndicators = [];
                        //определим индикаторы
                        var indicators = row.Cells[TableColIndexCompetenceIndicators].GetText(",")
                            .Split([',','\n'], options: StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                        foreach (var indicator in indicators) {
                            if (CompetenceAchievement.TryParseIndicator(indicator, out var achi)) {
                                CompetenceIndicators.Add(achi.Code);
                            }
                            else {
                                fos.Errors.Add($"Таблица оценочного средства [{Type.GetDescription()}]: не удалось определить код индикатора достижений - {indicator}");
                            }
                        }
                        //определим список вопросов
                        var allText = row.Cells[TableColIndexItems].GetText("\r\n", applyTrim: true);
                        if (!string.IsNullOrEmpty(allText)) {

                            //var par = row.Cells[TableColIndexItems].Paragraphs.FirstOrDefault();
                            //par.InsertParagraphAfterSelf()
                            //var xml = par.Xml;

                            var items = allText.Split("\r\n");
                            if (Type == EEvaluationTool.Testing) {
                                TestingItems = [];
                                var idx = 0;
                                var testingItem = new TestingItem() { Answers = [], RightAnswerIndicies = [] };
                                //var answers = new List<string>();
                                while (idx < items.Length) {
                                    var line = items[idx];
                                    if (m_testingQuestionMark.IsMatch(line) || idx == items.Length - 1) {
                                        if (!string.IsNullOrEmpty(testingItem.Question)) {
                                            TestingItems.Add(testingItem);
                                        }
                                        testingItem = new() {
                                            Question = line,
                                            Answers = [],
                                            RightAnswerIndicies = []
                                        };
                                    }
                                    else {
                                        testingItem.Answers.Add(line);
                                    }

                                    idx++;
                                }
                                if (TestingItems.Count == 0) {
                                    XmlTestingItems = row.Cells[TableColIndexItems].Paragraphs.Select(p => p.Xml.ToString()).ToList();
                                    //ar
                                }
                            }
                            else {
                                ListItems = items.ToList();
                            }
                        }
                        else {
                            fos.Errors.Add($"Таблица оценочного средства [{Type.GetDescription()}]: список вопросов не определён");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Проверка переданной таблицы на таблицу описания оценочных средств
        /// </summary>
        /// <param name="table"></param>
        /// <param name="colIndicatorsIndex"></param>
        /// <returns></returns>
        public static bool TestTable(Table table, out int colIndicatorsIndex) {
            var result = false;
            colIndicatorsIndex = -1;

            if (table.RowCount > 0 && table.Rows[0].Cells.Count >= 3) {
                //далее поищем заголовок "коды..."
                for (var col = table.Rows[0].Cells.Count - 1; col >= 0; col--) {
                    var cellText = table.Rows[0].Cells[col].GetText();
                    if (!string.IsNullOrEmpty(cellText) && cellText.StartsWith("коды", StringComparison.CurrentCultureIgnoreCase)) {
                        result = true;
                        colIndicatorsIndex = col;
                    }
                }
            }

            return result;
        }
    }
}
