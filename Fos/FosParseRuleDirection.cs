using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class FosParseRuleDirection : IDocParseRule<Fos> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Inline;
        public string MultilineConcatValue { get; set; } = " ";
        public string PropertyName { get; set; } = null;    //чтобы применялся Action
        public Type PropertyType { get; set; } = null;
        public List<(Regex marker, int inlineGroupIdx)> StartMarkers { get; set; } = [
            (new(@"(\d{2}\s*\.\s*\d{2}\s*\.\s*\d{2})\s+(.*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase), 1)
        ];
        public List<Regex> StopMarkers { get; set; } = null;
        public char[] TrimChars { get; set; } = null; // [' ', '«', '»', '"', '“', '”'];
        public Action<Fos, Match, string, Paragraph> Action { get; set; } = (fos, match, text, par) => {
            fos.DirectionCode = string.Join("", match.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries));
            fos.DirectionName = match.Groups[2].Value.Trim(' ', '«', '»', '"', '“', '”');
        };
        public bool MultyApply { get; set; } = false;
        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
