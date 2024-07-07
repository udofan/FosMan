using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class RpdParseRuleNextDisciplines : IDocParseRule<Rpd> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Inline;
        public string MultilineConcatValue { get; set; } = "\r\n";
        public string PropertyName { get; set; } = null; // nameof(Rpd.NextDisciplines);
        public Type PropertyType { get; set; } = null; //typeof(Rpd).GetProperty(nameof(Rpd.NextDisciplines))?.PropertyType;
        public List<(Regex marker, int catchGroupIdx)> StartMarkers { get; set; } = [
            //последующих учебных дисциплин
            //изучении последующих профессиональных дисциплин
            (new(@"послед[^.]+(учеб)?[^.]+дисциплин(ы)?[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 3),
            //изучения следующих дисциплин
            (new(@"следующих\s+учебных\s+дисциплин[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            (new(@"следующих\s+дисциплин[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //служат основой для более глубокого восприятия таких дисциплин как
            (new(@"глуб[^.]+восприят[^.]+так[^.]+дисциплин как\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //В соединении с дисциплинами
            (new(@"в\s+соединен[^.]+\s+дисциплинами\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //необходимы для освоения таких предметов как 
            (new(@"для\s+освоен[^.]+\s+предмет[^.]+\s+как\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //является базовым для последующего освоения программного материала ряда дисциплин
            (new(@"базов[^.]+для\s+последующ[^.]+\s+освоен[^.]+материала\s+ряда\s+дисциплин\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //является базовым для последующего освоения программного материала
            (new(@"базов[^.]+для\s+последующ[^.]+\s+освоен[^.]+материала\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //более успешному изучению связанных с ней дисциплин:
            (new(@"успешн[^.]+изуч[^.]+связанн[^.]+дисциплин[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            ////Наименования последующих направлений:
            (new(@"наименов[^.]+послед[^.]+направлений[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //более успешному изучению таких связанных с ней дисциплин, как
            (new(@"изуч[^.]+таких[^.]+связанн[^.]+дисциплин[^.]+как\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //является предшествующей для изучения дисциплин
            (new(@"являет[^.]+предшествующ[^.]+изучен[^.]+дисциплин[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //является фундаментальной базой для дальнейшего изучения 
            (new(@"являет[^.]+базой[^.]+дальнейшего[^.]+изучения\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //служат основой для более глубокого восприятия таких дисциплин как
            (new(@"служат\s+основой\s+для\s+более\s+глубокого\s+восприятия\s+таких\s+дисциплин\s+как\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //знания могут быть использованы при изучении таких дисциплин, как: 
            (new(@"знания\s+могут\s+быть\s+использованы\s+при\s+изучении\s+таких\s+дисциплин\s+как[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
        ];
        public List<(Regex marker, int catchGroupIdx)> StopMarkers { get; set; } = null;
        public char[] TrimChars { get; set; } = null; // [' ', '«', '»', '"', '“', '”'];
        public Action<DocParseRuleActionArgs<Rpd>> Action { get; set; } = args => {
            args.Target.NextDisciplines = args.Value;
            args.Target.FullTextNextDisciplines = args.Text;
        };
        public bool MultyApply { get; set; } = false;
        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
