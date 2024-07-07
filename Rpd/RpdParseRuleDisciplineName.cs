using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class RpdParseRuleDisciplineName : IDocParseRule<Rpd> {
        public EParseType Type { get; set; } = EParseType.Multiline;
        public string MultilineConcatValue { get; set; } = " ";
        public string PropertyName { get; set; } = nameof(Rpd.DisciplineName);
        public Type PropertyType { get; set; } = typeof(Rpd).GetProperty(nameof(Rpd.DisciplineName))?.PropertyType;
        public List<(Regex marker, int catchGroupIdx)> StartMarkers { get; set; } = [
            (new(@"рабочая\s+программа\s+дисциплины", RegexOptions.Compiled | RegexOptions.IgnoreCase), 1)
        ];
        public List<(Regex marker, int catchGroupIdx)> StopMarkers { get; set; } = [
            (new(@"^$", RegexOptions.Compiled | RegexOptions.IgnoreCase), -1),   //пустая строка
            (new(@"^([^»”]+)[»”]{1}", RegexOptions.Compiled), 1)                 //закрывающая кавычка
        ];
        public char[] TrimChars { get; set; } = [' ', '«', '»', '"', '“', '”'];
        public Action<DocParseRuleActionArgs<Rpd>> Action { get; set; } = null;
        public bool MultyApply { get; set; } = false;

        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
