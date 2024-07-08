using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;
using static FosMan.Enums;

namespace FosMan {
    internal class RpdParseRuleFormsOfStudyMultiline : IDocParseRule<Rpd> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Multiline;
        public string MultilineConcatValue { get; set; } = " ";
        public string PropertyName { get; set; } = null;    //чтобы применялся Action
        public Type PropertyType { get; set; } = null;
        public List<(Regex marker, int catchGroupIdx)> StartMarkers { get; set; } = [
            (new(@"^Форм\S+\s+обучения[:]*\s*$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 1)
        ];
        public List<(Regex marker, int catchGroupIdx)> StopMarkers { get; set; } = [
            (new(@"^$", RegexOptions.Compiled | RegexOptions.IgnoreCase), -1) //пустая строка
        ];
        public char[] TrimChars { get; set; } = null;
        public Action<DocParseRuleActionArgs<Rpd>> Action { get; set; } = (args) => {
            var rpd = args.Target;
            var items = args.Match.Groups[1].Value.Split(',', StringSplitOptions.TrimEntries);
            foreach (var item in items) {
                switch (item.Trim().ToLower()) {
                    case "очная":
                        rpd.FormsOfStudy.Add(EFormOfStudy.FullTime);
                        break;
                    case "заочная":
                        rpd.FormsOfStudy.Add(EFormOfStudy.PartTime);
                        break;
                    case "очно-заочная":
                        rpd.FormsOfStudy.Add(EFormOfStudy.MixedTime);
                        break;
                    default:
                        rpd.FormsOfStudy.Add(EFormOfStudy.Unknown);
                        break;
                }
            }
        };
        public bool MultyApply { get; set; } = false;

        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
