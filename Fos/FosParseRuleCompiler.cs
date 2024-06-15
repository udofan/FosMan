using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FosMan {
    internal class FosParseRuleCompiler : IDocParseRule<Fos> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Inline;
        public string MultilineConcatValue { get; set; } = string.Empty;
        public string Name { get; set; } = nameof(Fos.Compiler);
        public List<(Regex marker, int inlineGroupIdx)> StartMarkers { get; set; } = [
            (new(@"([А-Я]{1}\.\s*[А-Я]{1}\.\s*.+$)", RegexOptions.Compiled | RegexOptions.IgnoreCase), 1)
        ];
        public List<Regex> StopMarkers { get; set; } = null;
        public char[] TrimChars { get; set; } = null;
        public Action<Fos, Match, string> Action { get; set; } = null;

        //public bool Equals(IDocParseRule<Fos>? other) {
        //    return string.Compare(this.Name, other?.Name) == 0;
        //}

        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
