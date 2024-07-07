using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    /// <summary>
    /// НЕ ПРИМЕНЯТЬ - все в коде
    /// </summary>
    internal class RpdParseRuleSummary: IDocParseRule<Rpd> {
        public EParseType Type { get; set; } = EParseType.Inline;
        public string MultilineConcatValue { get; set; } = string.Empty;
        public string PropertyName { get; set; } = null;
        public Type PropertyType { get; set; } = null;
        public List<(Regex marker, int inlineGroupIdx)> StartMarkers { get; set; } = [
            //Основные задачи дисциплины
            (new(@"задачи\s+дисциплины[:]*.+$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
            //Основные задачи дисциплины
            (new(@"задачами\s+дисциплины\s+являю.+$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
        ];
        public List<Regex> StopMarkers { get; set; } = [
            new(@"^$", RegexOptions.Compiled | RegexOptions.IgnoreCase) //пустая строка
        ];
        public char[] TrimChars { get; set; } = null;
        public Action<DocParseRuleActionArgs<Rpd>> Action { get; set; } = args => {

        };
        public bool MultyApply { get; set; } = false;
        //public bool Equals(IDocParseRule<Fos>? other) {
        //    return string.Compare(this.Name, other?.Name) == 0;
        //}

        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
