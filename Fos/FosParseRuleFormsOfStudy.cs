using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static FosMan.Enums;

namespace FosMan {
    internal class FosParseRuleFormsOfStudy : IDocParseRule<Fos> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Inline;
        public string MultilineConcatValue { get; set; } = string.Empty;
        public string PropertyName { get; set; } = null;    //чтобы применялся Action
        public Type PropertyType { get; set; } = null;
        public List<(Regex marker, int inlineGroupIdx)> StartMarkers { get; set; } = [
            (new(@"Форм\S+\s+обучения[:]*\s+(.+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 1)
        ];
        public List<Regex> StopMarkers { get; set; } = null;
        public char[] TrimChars { get; set; } = null;
        public Action<Fos, Match, string> Action { get; set; } = (fos, match, text) => {
            var items = match.Groups[1].Value.Split(',', StringSplitOptions.TrimEntries);
            foreach (var item in items) {
                switch (item.Trim().ToLower()) {
                    case "очная":
                        fos.FormsOfStudy.Add(EFormOfStudy.FullTime);
                        break;
                    case "заочная":
                        fos.FormsOfStudy.Add(EFormOfStudy.PartTime);
                        break;
                    case "очно-заочная":
                        fos.FormsOfStudy.Add(EFormOfStudy.MixedTime);
                        break;
                    default:
                        fos.FormsOfStudy.Add(EFormOfStudy.Unknown);
                        break;
                }
            }
        };

        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
