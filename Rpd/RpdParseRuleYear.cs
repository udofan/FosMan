using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class RpdParseRuleYear : IDocParseRule<Rpd> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Inline;
        public string MultilineConcatValue { get; set; } = string.Empty;
        public string PropertyName { get; set; } = nameof(Rpd.Year);
        public Type PropertyType { get; set; } = typeof(Rpd).GetProperty(nameof(Rpd.Year))?.PropertyType;
        public List<(Regex marker, int catchGroupIdx)> StartMarkers { get; set; } = [
            (new(@"Москва\s+(\d{4})", RegexOptions.Compiled | RegexOptions.IgnoreCase), 1)
        ];
        public List<(Regex marker, int catchGroupIdx)> StopMarkers { get; set; } = null;
        public char[] TrimChars { get; set; } = null;
        public Action<DocParseRuleActionArgs<Rpd>> Action { get; set; } = null;
        public bool MultyApply { get; set; } = false;


        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
