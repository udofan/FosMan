using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class RpdParseRulePurpose : IDocParseRule<Rpd> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Multiline;
        public string MultilineConcatValue { get; set; } = "; ";
        public string PropertyName { get; set; } = nameof(Rpd.Purpose);
        public Type PropertyType { get; set; } = typeof(Rpd).GetProperty(nameof(Rpd.Purpose))?.PropertyType;
        public List<(Regex marker, int catchGroupIdx)> StartMarkers { get; set; } = [
            //Целью изучения дисциплины «Правоведение» является
            (new(@"целью\s+дисциплины\s+является\s+.+$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
            //Целью изучения дисциплины является
            (new(@"целью\s+изучения\s+дисциплины\s+является.+$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
            //Цель дисциплины -
            (new(@"цель\s+дисциплины[- ]+(.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
            //Цель изучения дисциплины
            (new(@"цель[ю]*\s+изучения\s+дисциплины[- ]+(.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
            //Цель изучения дисциплины заключается 
            (new(@"цель\s+изучения\s+дисциплины\s+заключ.+$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
            //Целями освоения учебной дисциплины 
            (new(@"целями\s+освоения\s+учебной\s+дисциплины.+$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
            //Целью дисциплины является
            (new(@"целью\s+дисциплины\s+является.+$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
            //Целью освоения дисциплины 
            (new(@"целью\s+освоения\s+дисциплины.+$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
            //Цели изучения дисциплины
            (new(@"цели\s+изучения\s+дисциплины.+$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 0),
        ];
        public List<(Regex marker, int catchGroupIdx)> StopMarkers { get; set; } = [
            (new(@"^$", RegexOptions.Compiled | RegexOptions.IgnoreCase), -1),      //пустая строка
            (new(@"задач", RegexOptions.Compiled | RegexOptions.IgnoreCase), -1),  //слово "задачи"
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
