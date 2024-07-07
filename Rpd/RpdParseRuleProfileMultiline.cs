using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class RpdParseRuleProfileMultiline : IDocParseRule<Rpd> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Multiline;
        public string MultilineConcatValue { get; set; } = " ";
        public string PropertyName { get; set; } = nameof(Rpd.Profile);
        public Type PropertyType { get; set; } = typeof(Rpd).GetProperty(nameof(Rpd.Profile))?.PropertyType;
        public List<(Regex marker, int inlineGroupIdx)> StartMarkers { get; set; } = [
            (new(@"(Профиль|Направленност[ь,и]\s+\S*\s*подготовки)[:]*\s*[«""“]*([^»""”]*)[»""”]*", RegexOptions.Compiled | RegexOptions.IgnoreCase), 1)
        ];
        public List<Regex> StopMarkers { get; set; } = [
            new(@"^$", RegexOptions.Compiled | RegexOptions.IgnoreCase),    //пустая строка
            new(@"[»”]{1}", RegexOptions.Compiled)                          //закрывающая кавычка
        ];
        public char[] TrimChars { get; set; } = [' ', '«', '»', '"', '“', '”'];
        public Action<DocParseRuleActionArgs<Rpd>> Action { get; set; } = null;
        public bool MultyApply { get; set; } = false;

        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
