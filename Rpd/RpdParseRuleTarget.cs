using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class RpdParseRuleTarget : IDocParseRule<Rpd> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Inline;
        public string MultilineConcatValue { get; set; } = string.Empty;
        public string PropertyName { get; set; } = nameof(Rpd.Target);
        public Type PropertyType { get; set; } = typeof(Rpd).GetProperty(nameof(Rpd.Target))?.PropertyType;
        public List<(Regex marker, int inlineGroupIdx)> StartMarkers { get; set; } = [
            //Целью изучения дисциплины «Правоведение» является
            (new(@"цель[^.]+дисциплины[^.]+является\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
            //Цель дисциплины -
            (new(@"цель\s+дисциплины[- ]+(.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0)
        ];
        public List<Regex> StopMarkers { get; set; } = null;
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
