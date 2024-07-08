using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class RpdParseRuleCompilerMultiline : IDocParseRule<Rpd> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Multiline;
        public string MultilineConcatValue { get; set; } = " ";
        public string PropertyName { get; set; } = nameof(Rpd.Compiler);
        public Type PropertyType { get; set; } = typeof(Rpd).GetProperty(nameof(Rpd.Compiler))?.PropertyType;
        public List<(Regex marker, int catchGroupIdx)> StartMarkers { get; set; } = [
            (new(@"Составитель:", RegexOptions.Compiled | RegexOptions.IgnoreCase), 1)
        ];
        public List<(Regex marker, int catchGroupIdx)> StopMarkers { get; set; } = [
            (new(@"^$", RegexOptions.Compiled | RegexOptions.IgnoreCase), -1)   //пустая строка
            //(new(@"^(.+)$", RegexOptions.Compiled))
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
