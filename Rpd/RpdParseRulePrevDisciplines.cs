using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class RpdParseRulePrevDisciplines : IDocParseRule<Rpd> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Inline;
        public string MultilineConcatValue { get; set; } = null;
        public string PropertyName { get; set; } = null; // nameof(Rpd.PrevDisciplines);
        public Type PropertyType { get; set; } = null; //typeof(Rpd).GetProperty(nameof(Rpd.PrevDisciplines))?.PropertyType;
        public List<(Regex marker, int catchGroupIdx)> StartMarkers { get; set; } = [
            //предварительное изучение следующих дисциплин
            (new(@"предварит[^.]*(изуч)?[^.]*(след)?[^.]+дисциплин[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 3),
            //следующих предшествующих дисциплин
            //формируемые предшествующими дисциплинами
            (new(@"предшест[^.]+дисциплин(ами|ы)?[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 2),
            //требуются знания таких дисциплин как 
            (new(@"треб[^.]+так[^.]+дисциплин\s+как\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //на которых базируется данная дисциплина
            (new(@"базир[^.]+данн[^.]+дисциплина[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //В круге главных источников, входящих в учебную дисциплину «Банковское де-ло» являются
            (new(@"глав[^.]+источник[^.]+входящ[^.]+дисципл[^.]+являются[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //базируется на знаниях и умениях, полученных обучающимися ранее в ходе освоения общеобразовательного программного материала по спряжённому курсу средней школы, а также ряда 
            (new(@"базир[^.]+на\s+знаниях[^.]+ранее[^.]+ряда\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //когда студенты ознакомлены с информатикой в части следующих направ-лений:
            (new(@"ознакомлен[^.]+части[^.]+следующ[^.]+направлений\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //логически связана с комплексом дисциплин:
            (new(@"логически\s+связана[^.]+комплекс[^.]+дисциплин[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //предшествуют следующие учебные курсы:
            (new(@"предшест[^.]+следующие[^.]+курсы[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 2),
            //Обязательным условием, обеспечивающим успешное освоение данной дис-циплины, являются хорошие знания обучающимися таких дисциплин, как 
            (new(@"хорош\s+знания[^.]+таких\s+дисциплин[^.]+как\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //дисциплина тесно связана с рядом общенаучных, экономических и специ-альных дисциплин, таких как 
            (new(@"дисциплин[^.]+связана[^.]+рядом[^.]+таких\s+как\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //когда студенты уже ознакомлены с дисциплиной 
            (new(@"уже[^.]+ознакомлены[^.]+диспиплин[^.]\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //формируемые предшествующими дисциплинами
            (new(@"формируем[^.]+предшествующ[^.]+дисциплинами[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //изучается параллельно с такими дисциплинами, как 
            (new(@"изучается[^.]+параллельно[^.]+такими\s+дисциплинами\s+как\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //..ранее изученных дисциплин:
            (new(@"ранее[^.]+изучен[^.]+дисциплин[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //полученных студентами при изучении дисциплин 
            (new(@"полученных[^.]+при[^.]+изучении[^.]+дисциплин\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //опирается на дисциплины, изучаемые студентом в магистратуре, такие как 
            (new(@"опирается\s+на\s+дисциплины[^.]+такие\s+как\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //сформированные у обучающихся в результате освоения дисциплин
            (new(@"сформированные\s+у\s+обучающихся\s+в\s+результате\s+освоения\s+дисциплин\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //полученных обучаемыми по дисциплинам
            (new(@"полученных\s+обучаемыми\s+по\s+дисциплинам\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //сформированных в ходе изучения дисциплин: 
            (new(@"сформированных\s+в\s+ходе\s+изучения\s+дисциплин[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //Дисциплина основывается на знании следующих дисциплин: 
            (new(@"основывается\s+на\s+знании\s+следующих\s+дисциплин[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
        ];
        public List<(Regex marker, int catchGroupIdx)> StopMarkers { get; set; } = null;
        public char[] TrimChars { get; set; } = null; // [' ', '«', '»', '"', '“', '”'];
        public Action<DocParseRuleActionArgs<Rpd>> Action { get; set; } = args => {
            args.Target.PrevDisciplines = args.Value;
            args.Target.FullTextPrevDisciplines = args.Text;
        };
        public bool MultyApply { get; set; } = false;
        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
