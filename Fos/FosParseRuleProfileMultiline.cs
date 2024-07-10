using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class FosParseRuleProfileMultiline : IDocParseRule<Fos> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Multiline;
        public string MultilineConcatValue { get; set; } = " ";
        public string PropertyName { get; set; } = nameof(Fos.Profile);
        public Type PropertyType { get; set; } = typeof(Fos).GetProperty(nameof(Fos.Profile))?.PropertyType;
        public List<(Regex marker, int catchGroupIdx)> StartMarkers { get; set; } = [
            (new(@"^профиль:$", RegexOptions.Compiled | RegexOptions.IgnoreCase), -1)
        ];
        public List<(Regex marker, int catchGroupIdx)> StopMarkers { get; set; } = [
            (new(@"^$", RegexOptions.Compiled | RegexOptions.IgnoreCase), -1) //пустая строка
        ];
        public char[] TrimChars { get; set; } = [' ', '«', '»', '"', '“', '”'];
        public Action<DocParseRuleActionArgs<Fos>> Action { get; set; } = null;
        public bool MultyApply { get; set; } = false;

        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
