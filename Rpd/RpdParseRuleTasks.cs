using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class RpdParseRuleTasks: IDocParseRule<Rpd> {
        public EParseType Type { get; set; } = EParseType.Inline;
        public string MultilineConcatValue { get; set; } = "\n";
        public string PropertyName { get; set; } = nameof(Rpd.Tasks);
        public Type PropertyType { get; set; } = typeof(Rpd).GetProperty(nameof(Rpd.Tasks))?.PropertyType;
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
        public Action<DocParseRuleActionArgs<Rpd>> Action { get; set; } = null;
        public bool MultyApply { get; set; } = false;
        //public bool Equals(IDocParseRule<Fos>? other) {
        //    return string.Compare(this.Name, other?.Name) == 0;
        //}

        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
