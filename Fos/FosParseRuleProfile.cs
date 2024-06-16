using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FosMan {
    internal class FosParseRuleProfile : IDocParseRule<Fos> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Multiline;
        public string MultilineConcatValue { get; set; } = " ";
        public string PropertyName { get; set; } = nameof(Fos.Profile);
        public Type PropertyType { get; set; } = typeof(Fos).GetProperty(nameof(Fos.Profile))?.PropertyType;
        public List<(Regex marker, int inlineGroupIdx)> StartMarkers { get; set; } = [
            (new(@"^профиль:$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 1)
        ];
        public List<Regex> StopMarkers { get; set; } = [
            new(@"^$", RegexOptions.Compiled | RegexOptions.IgnoreCase) //пустая строка
        ];
        public char[] TrimChars { get; set; } = [' ', '«', '»', '"', '“', '”'];
        public Action<Fos, Match, string> Action { get; set; } = null;

        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
