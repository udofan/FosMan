using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;
using static FosMan.Enums;

namespace FosMan {
    /// <summary>
    /// Выявление оценочных средств по заголовка "1. Эссе, реферат.", "2. Доклад." и т.п.
    /// </summary>
    internal class FosParseRuleEvaluationTools : IDocParseRule<Fos> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Inline;
        public string MultilineConcatValue { get; set; } = string.Empty;
        public string PropertyName { get; set; } = null;    //чтобы применялся Action
        public Type PropertyType { get; set; } = null;
        public List<(Regex marker, int inlineGroupIdx)> StartMarkers { get; set; } = [
            (new(@"^([\d+\.]*\d+)\.\s+(доклад)[\.]*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 2),
            (new(@"^([\d+\.]*\d+)\.\s+(опрос)[\.]*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 2),
            (new(@"^([\d+\.]*\d+)\.\s+(тестирование)[\.]*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 2),
            (new(@"^([\d+\.]*\d+)\.\s+(курсовая\s+работа)[\.]*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 2),
            (new(@"^([\d+\.]*\d+)\.\s+(практическая\s+работа)[\.]*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 2),
            (new(@"^([\d+\.]*\d+)\.\s+(контрольная\s+работа)[\.]*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 2),
            (new(@"^([\d+\.]*\d+)\.\s+(мини-кейсы)[\.]*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 2),
            (new(@"^([\d+\.]*\d+)\.\s+(деловая\s+игра)[\.]*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 2),
            (new(@"^([\d+\.]*\d+)\.\s+(эссе)[\.]*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 2),
            (new(@"^([\d+\.]*\d+)\.\s+(реферат)[\.]*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 2),
            (new(@"^([\d+\.]*\d+)\.\s+(эссе,\s+реферат)[\.]*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 2),
        ];
        public List<Regex> StopMarkers { get; set; } = null;
        public char[] TrimChars { get; set; } = null;
        public Action<Fos, Match, string, Paragraph> Action { get; set; } = (fos, match, text, par) => {
            fos.EvalTools ??= [];

            var items = match.Groups[2].Value.Split(',', StringSplitOptions.TrimEntries);
            foreach (var item in items) {
                if (EvalToolDic.TryGetValue(item.ToUpper(), out var evalTool)) {
                    //ищем таблицу ниже
                    var currPar = par;
                    for (var i = 0; i < 3; i++) {
                        if (currPar.FollowingTables?.Any() ?? false) {
                            var table = currPar.FollowingTables.FirstOrDefault();
                            if (table.RowCount > 0 && table.Rows[0].Cells.Count >= 3) {
                                //далее поищем заголовок "коды..."
                                for (var col = table.Rows[0].Cells.Count - 1; col >= 0; col--) {
                                    var cellText = table.Rows[0].Cells[col].GetText();
                                    if (!string.IsNullOrEmpty(cellText) && cellText.StartsWith("коды", StringComparison.CurrentCultureIgnoreCase)) {
                                        var tool = new EvaluationTool() {
                                            ChapterNum = match.Groups[1].Value,
                                            TableIndex = table.Index,
                                            Table = table,
                                            TableColIndexCompetenceIndicators = col,
                                            TableColIndexItems = col - 1,
                                            Type = evalTool
                                        };
                                        if (!fos.EvalTools.TryGetValue(evalTool, out var list)) {
                                            list = [];
                                            fos.EvalTools[evalTool] = list;
                                        }

                                        list.Add(tool);
                                        tool.ParseItems(fos);
                                        break;
                                    }
                                }
                            }
                        }
                        currPar = currPar.NextParagraph;
                    }
                }
            }
        };
        public bool MultyApply { get; set; } = true;

        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
