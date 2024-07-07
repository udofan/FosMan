using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Net.WebSockets;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;
using static FosMan.Enums;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace FosMan {
    internal class Rpd : BaseObj {
        //Кафедра
        static Regex m_regexDepartment = new(@"Кафедра\s+(.+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //Маркер конца титульной странцы "Москва 20xx"
        static Regex m_regexTitlePageEndMarker = new(@"Москва\s+(\d{4})", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //профиль/направленность подготовки
        static Regex m_regexProfileInline = new(@"(Профиль|Направленност[ь,и]\s+\S*\s*подготовки)[:]*\s+[«""“]*([^»""”]+)[»""”]*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexProfileMarker = new(@"(Профиль|Направленност[ь,и]\s+\S*\s*подготовки)[:]*\s*[«""“]*([^»""”]*)[»""”]*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //наименование в кавычках
        static Regex m_regexNameQuoted = new(@"[«""“]+(.+)[»""”]+$", RegexOptions.Compiled);
        //наименование в опциональных кавычках
        static Regex m_regexName = new(@"[«""“]*(.+)[»""”]*$", RegexOptions.Compiled);
        //направление подготовки
        static Regex m_regexDirection = new(@"(\d{2}\s*\.\s*\d{2}\s*\.\s*\d{2})\s+(.*)$", RegexOptions.Compiled);
        //форма обучения
        static Regex m_regexFormsOfStudy = new(@"Форм\S+\s+обучения[:]*\s+(.+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //составитель
        static Regex m_regexCompilerInline = new(@"Составитель[:]*\s+(.+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexCompilerMarker = new(@"Составитель:", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //маркер [РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ]
        static Regex m_regexRpdMarker = new(@"рабочая\s+программа\s+дисциплины", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static List<(Regex, int)> m_regexPrevDisciplines = new() {
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
        };
        static List<(Regex, int)> m_regexNextDisciplines = new() {
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
        };
        static List<(Regex, int)> m_regexTarget = new() {
            //Целью изучения дисциплины «Правоведение» является
            (new(@"цель[^.]+дисциплины[^.]+является\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1),
            //Цель дисциплины -
            (new(@"цель\s+дисциплины[- ]+(.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1)
        };
        static List<(Regex, int)> m_regexTasks = new() {
            //Основные задачи дисциплины
            (new(@"задачи[^.]+дисциплины[:]*", RegexOptions.IgnoreCase | RegexOptions.Compiled), 1)
        };
        static List<Regex> m_regexSummaryBeginMarkers = new() {
            new(@"^Краткое\s+содержание\s+дисциплины[:]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            //Краткое содержание дисциплины (по разделам и темам):
            //Краткое содержание дисциплины (по расширенным темам):
            new(@"^Краткое\s+содержание\s+дисциплины\s+\(по(.+)темам\)[:.]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new(@"^Содержание\s+дисциплины,\s+структурированное\s+по\s+разделам\s+\(темам\)[:.]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new(@"^Содержание\s+разделов\s+и\s+тем[:.]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            //Содержание разделов и тем дисциплины.
            new(@"^Содержание\s+разделов\s+и\s+тем\s+дисциплины[:.]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            //new(@"^Содержание\s+дисциплины,\s+структурированное\s+по\s+темам[:.]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            //Содержание дисциплины, структурированное по разделам
            new(@"^Содержание\s+дисциплины,\s+структурированное\s+по\s+(темам|разделам)[:.]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new(@"^Тематическое\s+содержание\s+разделов\s+дисциплины[:.]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new(@"^Тематическое\s+содержание\s+дисциплины[:.]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new(@"^Содержание\s+разделов\s+дисциплины[:.]*$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new(@"^Содержание\s+дисциплины:$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            new(@"^Краткое\s+содержание\s+разделов\s+дисциплины(.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            //Содержание и структура дисциплины «Элективного курса по физической 
            new(@"^Содержание\s+и\s+структура\s+дисциплины(.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled)
            
        };
        static Regex m_regexNumberedHeader = new(@"^(\d+)\.(.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        static List<Regex> m_regexQuestionListBeginMarker = new() {
            //Примерные вопросы к экзамену
            new(@"примерные\s+вопросы\s+к\s+(зачету|экзамену)[:.]*", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            //Вопросы для подготовки к зачету:
            new(@"Вопросы\s+для\s+подготовки\s+к\s+(зачету|экзамену)[:.]*", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            //Теоретический блок вопросов
            new(@"Теоретический\s+блок\s+вопросов[:]*", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            //Вопросы к зачету
            new(@"вопросы\s+к\s+(зачету|экзамену)[:]*", RegexOptions.IgnoreCase | RegexOptions.Compiled)
        };
        static List<Regex> m_regexReferencesBase = new() {
            new(@"основная.+литература", RegexOptions.IgnoreCase | RegexOptions.Compiled)
        };
        static List<Regex> m_regexReferencesExtra = new() {
            new(@"дополнительная.+литература", RegexOptions.IgnoreCase | RegexOptions.Compiled)
        };
        /// <summary>
        /// Ключ (по нему сравниваются РПД и ФОС)
        /// </summary>
        [JsonIgnore]
        public string Key { get => App.NormalizeName($"{DirectionCode}_{Profile}_{DisciplineName}"); }
        /// <summary>
        /// Год
        /// </summary>
        [JsonInclude]
        public string Year { get; set; }
        /// <summary>
        /// Дисциплина
        /// </summary>
        [JsonInclude]
        public string DisciplineName { get; set; }
        /// <summary>
        /// Профиль
        /// </summary>
        [JsonInclude]
        public string Profile { get; set; }
        /// <summary>
        /// Кафедра
        /// </summary>
        [JsonInclude]
        public string Department { get; set; }
        /// <summary>
        /// Код направления подготовки
        /// </summary>
        [JsonInclude]
        public string DirectionCode { get; set; }
        /// <summary>
        /// Направление подготовки
        /// </summary>
        [JsonInclude]
        public string DirectionName { get; set; }
        /// <summary>
        /// Формы обучения
        /// </summary>
        [JsonInclude]
        public List<EFormOfStudy> FormsOfStudy { get; set; }
        /// <summary>
        /// Матрица компетенций (п. 1 в РПД)
        /// </summary>
        [JsonInclude]
        public CompetenceMatrix CompetenceMatrix { get; set; }
        /// <summary>
        /// Исходный docx-файл
        /// </summary>
        [JsonInclude]
        public string SourceFileName { get; set; }
        /// <summary>
        /// Учебная работа по формам обучения
        /// </summary>
        [JsonInclude]
        public Dictionary<EFormOfStudy, EducationalWork> EducationalWorks { get; set; }
        /// <summary>
        /// Цель дисциплины
        /// </summary>
        [JsonInclude]
        public string Target { get; set; }
        /// <summary>
        /// Задачи дисциплины
        /// </summary>
        [JsonInclude]
        public string Tasks { get; set; }
        /// <summary>
        /// Выявленные ошибки
        /// </summary>
        [JsonIgnore]
        public ErrorList Errors { get; set; }
        /// <summary>
        /// Список доп. ошибок, выявленных при проверке
        /// </summary>
        [JsonIgnore]
        public ErrorList ExtraErrors { get; set; }
        /// <summary>
        /// Составитель
        /// </summary>
        [JsonInclude]
        public string Compiler { get; set; }
        /// <summary>
        /// Предшествующие дисциплины
        /// </summary>
        [JsonInclude]
        public string PrevDisciplines { get; set; }
        /// <summary>
        /// Последующие дисциплины
        /// </summary>
        [JsonInclude]
        public string NextDisciplines { get; set; }
        /// <summary>
        /// Параграфы содержания разделов и тем
        /// </summary>
        [JsonIgnore]
        public List<Paragraph> SummaryParagraphs { get; set; }
        /// <summary>
        /// Параграфы вопросов к экзамену/зачету
        /// </summary>
        [JsonInclude]
        public List<string> QuestionList { get; set; }
        /// <summary>
        /// Список базовых источников
        /// </summary>
        [JsonInclude]
        public List<string> ReferencesBase { get; set; }
        /// <summary>
        /// Список доп. источников
        /// </summary>
        [JsonInclude]
        public List<string> ReferencesExtra { get; set; }

        /// <summary>
        /// Очистка РПД
        /// </summary>
        public void Clear() {
            this.Errors = new();
            this.ExtraErrors = new();
            this.EducationalWorks = [];
            this.FormsOfStudy = [];
            this.SummaryParagraphs = [];
            this.QuestionList = [];
            this.ReferencesBase = [];
            this.ReferencesExtra = [];
            this.CompetenceMatrix = null;
            this.Compiler = "";
            this.Department = "";
            this.DirectionCode = "";
            this.DirectionName = "";
            this.DisciplineName = "";
            this.NextDisciplines = "";
            this.PrevDisciplines = "";
            this.Profile = "";
            this.Year = "";
        }

        /// <summary>
        /// Загрузка РПД из файла (старая версия)
        /// </summary>
        /// <param name="fileName"></param>
        public static Rpd LoadFromFileOld(string fileName, Rpd rpd = null) {
            rpd ??= new();
            rpd.Clear();
            rpd.SourceFileName = fileName;

            try {
                using (var docx = DocX.Load(fileName)) {
                    //парсинг титульной страницы
                    var parIdx = 0;
                    var disciplineTestReady = false;    //флаг готовности поиска имени дисциплины
                    var disciplineName = "";
                    var profileTestReady = false;
                    var profileField = "";              //если профиль разнесен по неск. строкам, здесь он будет накапливаться
                    var compilerTestReady = false;
                    //var fullTimeTestTableReady = false;
                    //var mixedTimeTestTableReady = false;
                    List<Paragraph> parList = [];
                    List<string> textList = [];
                    var collectPars = false;
                    var collectText = false;
                    var collectSummary = false;
                    var collectQuestions = false;
                    var collectBaseReferences = false;
                    var collectExtraReferences = false;

                    foreach (var par in docx.Paragraphs) {
                        var text = par.Text.Trim(' ', '\r', '\n', '\t');
                        //если в предыд. итерацию был обнаружен маркер [РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ]
                        if (disciplineTestReady) {
                            if (string.IsNullOrEmpty(text)) {
                                var matchDiscipline = m_regexName.Match(disciplineName);
                                if (matchDiscipline.Success) {
                                    rpd.DisciplineName = matchDiscipline.Groups[1].Value.Trim(' ', '«', '»', '"', '“', '”');
                                    disciplineTestReady = false;
                                }
                            }
                            else {
                                if (disciplineName.Length > 0) disciplineName += " ";
                                disciplineName += text;                                //проверка на дисциплину с кавычками
                                var matchDiscipline = m_regexNameQuoted.Match(disciplineName);
                                if (matchDiscipline.Success) {
                                    rpd.DisciplineName = matchDiscipline.Groups[1].Value.Trim(' ', '«', '»', '"', '“', '”');
                                    disciplineTestReady = false;
                                }
                                continue;
                            }
                        }

                        if (string.IsNullOrWhiteSpace(text)) {
                            continue;
                        }
                        //кафедра
                        if (string.IsNullOrEmpty(rpd.Department)) {
                            var matchDepartment = m_regexDepartment.Match(text);
                            if (matchDepartment.Success) rpd.Department = matchDepartment.Groups[1].Value.Trim(' ', '«', '»', '"', '“', '”');
                        }
                        //предшествующие дисциплины
                        if (string.IsNullOrEmpty(rpd.PrevDisciplines)) {
                            (Match match, int idx) matchedItem = m_regexPrevDisciplines.Select(x => (x.Item1.Match(text), x.Item2)).FirstOrDefault(x => x.Item1.Success);
                            if (matchedItem.match != null) {
                                rpd.PrevDisciplines = matchedItem.match.Groups[matchedItem.idx].Value.Trim(' ');
                            }
                        }
                        //последующие дисциплины
                        if (string.IsNullOrEmpty(rpd.NextDisciplines)) {
                            (Match match, int idx) matchedItem = m_regexNextDisciplines.Select(x => (x.Item1.Match(text), x.Item2)).FirstOrDefault(x => x.Item1.Success);
                            if (matchedItem.match != null) {
                                rpd.NextDisciplines = matchedItem.match.Groups[matchedItem.idx].Value.Trim(' ');
                            }
                        }
                        //профиль
                        if (string.IsNullOrEmpty(rpd.Profile)) {
                            if (profileTestReady) {
                                if (profileField.Length > 0) profileField += " ";
                                profileField += text;

                                var matchProfile = m_regexName.Match(profileField);
                                if (matchProfile.Success) {
                                    rpd.Profile = matchProfile.Groups[1].Value.Trim(' ', '«', '»', '"', '“', '”');
                                    profileTestReady = false;
                                }
                            }
                            else {
                                var matchProfile = m_regexProfileInline.Match(text);
                                if (matchProfile.Success) {
                                    rpd.Profile = matchProfile.Groups[2].Value.Trim(' ', '«', '»', '"', '“', '”');
                                }
                                else {
                                    var matchProfileMarker = m_regexProfileMarker.Match(text);
                                    if (matchProfileMarker.Success) {
                                        if (profileField.Length > 0) profileField += " ";
                                        profileField += matchProfileMarker.Groups[2].Value;
                                        profileTestReady = true;
                                    }
                                    //profileTestReady = matchProfileMarker.Success;
                                }
                            }
                        }
                        //составитель
                        if (string.IsNullOrEmpty(rpd.Compiler)) {
                            if (compilerTestReady) {
                                var matchCompiler = m_regexName.Match(text);
                                if (matchCompiler.Success) {
                                    rpd.Compiler = matchCompiler.Groups[1].Value.Trim(' ', '«', '»', '"', '“', '”');
                                    compilerTestReady = false;
                                }
                            }
                            else {
                                var matchCompiler = m_regexCompilerInline.Match(text);
                                if (matchCompiler.Success) {
                                    rpd.Compiler = matchCompiler.Groups[2].Value.Trim(' ', '«', '»', '"', '“', '”');
                                }
                                else {
                                    var matchCompilerMarker = m_regexCompilerMarker.Match(text);
                                    compilerTestReady = matchCompilerMarker.Success;
                                }
                            }
                        }
                        //сбор параграфов
                        if (collectPars) {
                            var stopCollecting = false;
                            if (collectSummary) {
                                var m = m_regexNumberedHeader.Match(text);
                                stopCollecting = m.Success && int.Parse(m.Groups[1].Value) >= 4 && //номер раздела как правило 5
                                                 (m.Groups[2].Value.Contains("учебно-методич", StringComparison.CurrentCultureIgnoreCase) || //в нем есть слово "методич"
                                                 par.MagicText.FirstOrDefault().formatting?.Bold == true); /* или он жирный */
                                if (!stopCollecting) {
                                    if (//par.MagicText.FirstOrDefault().formatting?.Bold == true && 
                                        text.Contains("учебно-методич", StringComparison.CurrentCultureIgnoreCase) &&
                                        par.IsListItem) {
                                        stopCollecting = true;
                                    }
                                }
                            }

                            if (stopCollecting) {
                                if (collectSummary) rpd.SummaryParagraphs = parList;
                                collectPars = false;
                            }
                            else {
                                parList.Add(par);
                            }
                        }
                        //сбор текстовых значений
                        if (collectText) {
                            var stopCollecting = false;
                            var clearedText = text;
                            var matchNumberedItem = m_regexNumberedHeader.Match(text);
                            //if (collectQuestions) {
                            //    stopCollecting = matchNumberedItem.Success && int.Parse(matchNumberedItem.Groups[1].Value) >= 6 && //номер раздела как правило 8
                            //                    (matchNumberedItem.Groups[2].Value.Contains("Перечень", StringComparison.CurrentCultureIgnoreCase) || //в нем есть слово "Перечень"
                            //                    par.MagicText.FirstOrDefault().formatting?.Bold == true); /* или он жирный */
                            //}
                            if (!stopCollecting && collectBaseReferences) { //основная литература - конец списка на заголовке доп. литературы?
                                stopCollecting = m_regexReferencesExtra.Any(r => r.IsMatch(text));
                            }
                            if (!stopCollecting && collectExtraReferences) { //доп. лит-ра - конец списка?
                                stopCollecting = matchNumberedItem.Success && int.Parse(matchNumberedItem.Groups[1].Value) >= 7 && //номер раздела как правило 8
                                                 (matchNumberedItem.Groups[2].Value.Contains("Профессиональные", StringComparison.CurrentCultureIgnoreCase) || //в нем есть слово "Профессиональные"
                                                 par.MagicText.FirstOrDefault().formatting?.Bold == true); /* или он жирный */
                            }
                            if (!stopCollecting) {
                                if (matchNumberedItem.Success) {
                                    clearedText = matchNumberedItem.Groups[2].Value.Trim(' ', '\r', '\n', '\t');
                                    var num = int.Parse(matchNumberedItem.Groups[1].Value);
                                    if (num < textList.Count) {
                                        stopCollecting = true;
                                    }
                                }
                                else { //не подошло по выражению - проверим как элемент нумерованного списка
                                    if (par.IsListItem && par.ListItemType == ListItemType.Numbered) {
                                        if (int.TryParse(par.GetListItemNumber(), out var num)) {
                                            if (num < textList.Count) {
                                                stopCollecting = true;
                                            }
                                        }
                                    }
                                    //else {
                                    //    stopCollecting = true;
                                    //}
                                }
                            }
                            if (stopCollecting) {
                                if (collectQuestions) rpd.QuestionList = textList.ToList();
                                if (collectBaseReferences) rpd.ReferencesBase = textList.ToList();
                                if (collectExtraReferences) rpd.ReferencesExtra = textList.ToList();
                                collectText = false;
                                collectQuestions = false;
                                collectBaseReferences = false;
                                collectExtraReferences = false;
                            }
                            else {
                                textList.Add(clearedText); 
                            }
                            //}
                        }
                        //маркер начала содержания разделов и тем
                        if (rpd.SummaryParagraphs.Count == 0) {
                            if (m_regexSummaryBeginMarkers.Any(r => r.IsMatch(text))) {
                                parList.Clear();
                                collectPars = true;
                                collectSummary = true;
                            }
                        }
                        //маркер вопросов к зкз/зачету
                        if (rpd.QuestionList.Count == 0) {
                            if (m_regexQuestionListBeginMarker.Any(r => r.IsMatch(text))) {
                                textList.Clear();
                                collectText = true;
                                collectQuestions = true;
                            }
                        }
                        //маркер основной литературы
                        if (rpd.ReferencesBase.Count == 0) {
                            if (m_regexReferencesBase.Any(r => r.IsMatch(text))) {
                                if (collectQuestions) {
                                    rpd.QuestionList = textList.ToList();
                                    collectQuestions = false;
                                }
                                textList.Clear();
                                collectText = true;
                                collectBaseReferences = true;
                            }
                        }
                        //маркер доп. литературы
                        if (rpd.ReferencesExtra.Count == 0) {
                            if (m_regexReferencesExtra.Any(r => r.IsMatch(text))) {
                                textList.Clear();
                                collectText = true;
                                collectExtraReferences = true;
                            }
                        }
                        //направление подготовки
                        if (string.IsNullOrEmpty(rpd.DirectionCode)) {
                            var matchDirection = m_regexDirection.Match(text);
                            if (matchDirection.Success) {
                                rpd.DirectionCode = string.Join("", matchDirection.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries));
                                rpd.DirectionName = matchDirection.Groups[2].Value.Trim(' ', '«', '»', '"', '“', '”');
                            }
                        }
                        //формы обучения
                        //вариант 1: в двух абзацах
                        if (text.Equals("форма обучения", StringComparison.CurrentCultureIgnoreCase)) {
                            text += $" {par.NextParagraph.Text}";
                        }
                        //вариант 2: в одной строке
                        if (rpd.FormsOfStudy.Count == 0) {
                            var matchFormsOfStudy = m_regexFormsOfStudy.Match(text);
                            if (matchFormsOfStudy.Success) {
                                var items = matchFormsOfStudy.Groups[1].Value.Split(',', StringSplitOptions.TrimEntries);
                                foreach (var item in items) {
                                    switch (item.Trim().ToLower()) {
                                        case "очная":
                                            rpd.FormsOfStudy.Add(EFormOfStudy.FullTime);
                                            break;
                                        case "заочная":
                                            rpd.FormsOfStudy.Add(EFormOfStudy.PartTime);
                                            break;
                                        case "очно-заочная":
                                            rpd.FormsOfStudy.Add(EFormOfStudy.MixedTime);
                                            break;
                                        default:
                                            rpd.FormsOfStudy.Add(EFormOfStudy.Unknown);
                                            break;

                                    }
                                }
                            }
                        }

                        //проверка на маркер [РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ]
                        if (string.IsNullOrEmpty(rpd.DisciplineName)) {
                            var matchRpdMarker = m_regexRpdMarker.Match(text);
                            disciplineTestReady = matchRpdMarker.Success;
                        }
                        //если встретился маркер конца титульной страницы
                        var matchPageEnd = m_regexTitlePageEndMarker.Match(text);
                        if (matchPageEnd.Success) {
                            rpd.Year = matchPageEnd.Groups[1].Value;
                        }
                        parIdx++;
                    }

                    var fullTimeTableIsOk = false;
                    var mixedTimeTableIsOk = false;
                    var partTimeTableIsOk = false;

                    //функция для нормализации текста: убираем лишние пробелы и приведение к UpperCase
                    Func<string, string> normalizeText = t => string.Join(" ", t.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)).ToUpper();

                    //проверка таблиц
                    foreach (var table in docx.Tables) {
                        var keepTestTable = true;
                        if (!rpd.EducationalWorks.Any()) {
                            keepTestTable = !App.TestForSummaryTableForEducationalWorks(table, rpd.EducationalWorks, PropertyAccess.Get);
                        }
                        if (keepTestTable && rpd.CompetenceMatrix == null) {
                            if (CompetenceMatrix.TestTable(table, out var format) && format == ECompetenceMatrixFormat.Rpd) {
                                if (App.TestForTableOfCompetenceMatrix(table, format, out var matrix, out var errors)) {
                                    rpd.CompetenceMatrix = matrix;
                                    keepTestTable = false;
                                }
                                if (errors.Any()) rpd.Errors.AddRange(errors);
                            }
                            //testTable = !App.TestForTableOfCompetenceMatrix(table, rpd);
                        }
                        if (keepTestTable) {
                            EEvaluationTool[] evalTools = null;
                            string[][] studyResults = null;
                            if (App.TestForEduWorkTable(docx, table, rpd, PropertyAccess.Get, EEduWorkFixType.All, ref evalTools, ref studyResults, 
                                                        0, null, null, /* параметры для режима исправления - здесь они не нужны */
                                                        out var formOfStudy)) {
                                //получим список модулей обучения с оценочными средствами и компетенциями
                                var eduWork = rpd.EducationalWorks[formOfStudy];

                                eduWork.Modules = [];
                                //var evalToolDic = Enum.GetValues(typeof(EEvaluationTool)).Cast<EEvaluationTool>().ToDictionary(x => x.GetDescription().ToUpper(), x => x);

                                for (var row = eduWork.TableTopicStartRow; row <= eduWork.TableTopicLastRow; row++) { //минус ряд с "зачетом", минус ряд с "итого"
                                    var module = new StudyModule();
                                    var col = eduWork.TableColTopic; // startNumCol - 1;
                                    var cellIdx = col;
                                    while (col < eduWork.TableMaxColCount) {
                                        if (col == eduWork.TableColTopic) {
                                            if (table.Rows[row].Cells.Count <= cellIdx) {
                                                var tt = 0;
                                            }
                                            try {
                                                module.Topic = table.Rows[row].Cells[cellIdx].GetText();
                                            }
                                            catch (Exception exx) {
                                                var tt = 0;
                                            }
                                        }
                                        if (col == eduWork.TableColEvalTools) { //оценочные средства
                                            module.EvaluationTools = [];
                                            //var counts = table.Rows.Select(r => r.Cells.Count).ToList();
                                            //module.Topic = table.Rows[row].Cells[col].GetText();
                                            foreach (var t in table.Rows[row].Cells[cellIdx].GetText().Split(',', '\n', ';')) {
                                                //убираем лишние пробелы
                                                if (Enums.EvalToolDic.TryGetValue(normalizeText(t), out var tool)) {
                                                    //var normalizedText = string.Join(" ", t.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)).ToUpper();
                                                    //if (Enum.TryParse(normalizedText, true, out EEvaluationTool tool)) {
                                                    module.EvaluationTools.Add(tool);
                                                }
                                                else {
                                                    if (string.IsNullOrEmpty(t)) {
                                                        rpd.Errors.Add(EErrorType.RpdMissingEvalTool, $"Тема [{module.Topic}] ({formOfStudy.GetDescription()})");
                                                    }
                                                    else {
                                                        rpd.Errors.Add(EErrorType.RpdUnknownEvalTool, $"Значение - {t}, тема [{module.Topic}] ({formOfStudy.GetDescription()})");
                                                    }
                                                }
                                            }
                                        }
                                        if (col == eduWork.TableColCompetenceResults) { //результаты обучения - компетенции
                                            module.CompetenceResultCodes = table.Rows[row].Cells[cellIdx].GetText(",").Split(',', '\n', ';').ToHashSet();
                                        }
                                        col += table.Rows[row].Cells[cellIdx].GridSpan + 1; //переход на след. ячейку с учетом объединений
                                        cellIdx++;                                          //индексы ячеек последовательны
                                    }
                                    eduWork.Modules.Add(module);
                                }

                                keepTestTable = false;  //таблица обработана
                            }
                        }

                    }
                    //итоговая проверка
                    if (rpd.FormsOfStudy.Count != rpd.EducationalWorks.Count) {
                        foreach (var form in rpd.FormsOfStudy) {
                            if (!rpd.EducationalWorks.ContainsKey(form)) {
                                rpd.Errors.Add(EErrorType.RpdMissingEduWork, $"Форма обучения [{form}]");
                            }
                        }
                    }
                    if (rpd.CompetenceMatrix == null || !rpd.CompetenceMatrix.IsLoaded) {
                        rpd.Errors.Add(EErrorType.RpdMissingCompetenceMatrix);
                    }
                    if (rpd.SummaryParagraphs.Count == 0) {
                        rpd.Errors.Add(EErrorType.RpdMissingTOC);
                    }
                    if (rpd.QuestionList.Count == 0) {
                        rpd.Errors.Add(EErrorType.RpdMissingQuestionList);
                    }
                    if (rpd.EducationalWorks.Count != rpd.FormsOfStudy.Count) {
                        rpd.Errors.Add(EErrorType.RpdNotFullEduWorks, $"(только для: {string.Join(", ", rpd.EducationalWorks?.Select(x => x.Key.GetDescription()))})");
                    }
                    if (rpd.ReferencesBase.Count == 0) {
                        rpd.Errors.Add(EErrorType.RpdMissingReferencesBase);
                    }
                    if (rpd.ReferencesExtra.Count == 0) {
                        rpd.Errors.Add(EErrorType.RpdMissingReferencesExtra);
                    }
                    foreach (var eduWork in rpd.EducationalWorks) {
                        if (eduWork.Value.Table == null) {
                            rpd.Errors.Add(EErrorType.RpdMissingEduWorkTable, $"(форма обучения [{eduWork.Key.GetDescription()}])");
                        }
                    }
                    if (string.IsNullOrEmpty(rpd.Department)) rpd.Errors.Add(EErrorType.RpdMissingDepartment);
                    if (string.IsNullOrEmpty(rpd.Profile)) rpd.Errors.Add(EErrorType.RpdMissingProfile);
                    if (string.IsNullOrEmpty(rpd.Year)) rpd.Errors.Add(EErrorType.RpdMissingYear);
                    if (string.IsNullOrEmpty(rpd.DirectionCode)) rpd.Errors.Add(EErrorType.RpdMissingDirectionCode);
                    if (string.IsNullOrEmpty(rpd.DirectionName)) rpd.Errors.Add(EErrorType.RpdMissingDirectionName);
                    if (string.IsNullOrEmpty(rpd.DisciplineName)) rpd.Errors.Add(EErrorType.RpdMissingDisciplineName);
                }
            }
            catch (Exception ex) {
                rpd.Errors.AddException(ex);
            }

            return rpd;
        }

        /// <summary>
        /// Загрузка РПД из файла
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="rpd"></param>
        /// <returns></returns>
        public static Rpd LoadFromFile(string fileName, Rpd rpd = null) {
            rpd ??= new();
            rpd.Clear();
            rpd.SourceFileName = fileName;

            try {
                if (!DocParser.TryParse(fileName, rpd, RpdParser.Rules, out var errorList)) {
                    rpd.Errors.AddRange(errorList.Items);
                    return null;
                }
                else {
                    using (var docx = DocX.Load(fileName)) {
                        //парсинг титульной страницы
                        var parIdx = 0;
                        var disciplineTestReady = false;    //флаг готовности поиска имени дисциплины
                        var disciplineName = "";
                        var profileTestReady = false;
                        var profileField = "";              //если профиль разнесен по неск. строкам, здесь он будет накапливаться
                        var compilerTestReady = false;
                        //var fullTimeTestTableReady = false;
                        //var mixedTimeTestTableReady = false;
                        List<Paragraph> parList = [];
                        List<string> textList = [];
                        var collectPars = false;
                        var collectText = false;
                        var collectSummary = false;
                        var collectQuestions = false;
                        var collectBaseReferences = false;
                        var collectExtraReferences = false;

                        foreach (var par in docx.Paragraphs) {
                            var text = par.Text.Trim(' ', '\r', '\n', '\t');
                            //если в предыд. итерацию был обнаружен маркер [РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ]
                            if (disciplineTestReady) {
                                if (string.IsNullOrEmpty(text)) {
                                    var matchDiscipline = m_regexName.Match(disciplineName);
                                    if (matchDiscipline.Success) {
                                        rpd.DisciplineName = matchDiscipline.Groups[1].Value.Trim(' ', '«', '»', '"', '“', '”');
                                        disciplineTestReady = false;
                                    }
                                }
                                else {
                                    if (disciplineName.Length > 0) disciplineName += " ";
                                    disciplineName += text;                                //проверка на дисциплину с кавычками
                                    var matchDiscipline = m_regexNameQuoted.Match(disciplineName);
                                    if (matchDiscipline.Success) {
                                        rpd.DisciplineName = matchDiscipline.Groups[1].Value.Trim(' ', '«', '»', '"', '“', '”');
                                        disciplineTestReady = false;
                                    }
                                    continue;
                                }
                            }

                            if (string.IsNullOrWhiteSpace(text)) {
                                continue;
                            }
                            //кафедра
                            if (string.IsNullOrEmpty(rpd.Department)) {
                                var matchDepartment = m_regexDepartment.Match(text);
                                if (matchDepartment.Success) rpd.Department = matchDepartment.Groups[1].Value.Trim(' ', '«', '»', '"', '“', '”');
                            }
                            //предшествующие дисциплины
                            if (string.IsNullOrEmpty(rpd.PrevDisciplines)) {
                                (Match match, int idx) matchedItem = m_regexPrevDisciplines.Select(x => (x.Item1.Match(text), x.Item2)).FirstOrDefault(x => x.Item1.Success);
                                if (matchedItem.match != null) {
                                    rpd.PrevDisciplines = matchedItem.match.Groups[matchedItem.idx].Value.Trim(' ');
                                }
                            }
                            //последующие дисциплины
                            if (string.IsNullOrEmpty(rpd.NextDisciplines)) {
                                (Match match, int idx) matchedItem = m_regexNextDisciplines.Select(x => (x.Item1.Match(text), x.Item2)).FirstOrDefault(x => x.Item1.Success);
                                if (matchedItem.match != null) {
                                    rpd.NextDisciplines = matchedItem.match.Groups[matchedItem.idx].Value.Trim(' ');
                                }
                            }
                            //профиль
                            if (string.IsNullOrEmpty(rpd.Profile)) {
                                if (profileTestReady) {
                                    if (profileField.Length > 0) profileField += " ";
                                    profileField += text;

                                    var matchProfile = m_regexName.Match(profileField);
                                    if (matchProfile.Success) {
                                        rpd.Profile = matchProfile.Groups[1].Value.Trim(' ', '«', '»', '"', '“', '”');
                                        profileTestReady = false;
                                    }
                                }
                                else {
                                    var matchProfile = m_regexProfileInline.Match(text);
                                    if (matchProfile.Success) {
                                        rpd.Profile = matchProfile.Groups[2].Value.Trim(' ', '«', '»', '"', '“', '”');
                                    }
                                    else {
                                        var matchProfileMarker = m_regexProfileMarker.Match(text);
                                        if (matchProfileMarker.Success) {
                                            if (profileField.Length > 0) profileField += " ";
                                            profileField += matchProfileMarker.Groups[2].Value;
                                            profileTestReady = true;
                                        }
                                        //profileTestReady = matchProfileMarker.Success;
                                    }
                                }
                            }
                            //составитель
                            if (string.IsNullOrEmpty(rpd.Compiler)) {
                                if (compilerTestReady) {
                                    var matchCompiler = m_regexName.Match(text);
                                    if (matchCompiler.Success) {
                                        rpd.Compiler = matchCompiler.Groups[1].Value.Trim(' ', '«', '»', '"', '“', '”');
                                        compilerTestReady = false;
                                    }
                                }
                                else {
                                    var matchCompiler = m_regexCompilerInline.Match(text);
                                    if (matchCompiler.Success) {
                                        rpd.Compiler = matchCompiler.Groups[2].Value.Trim(' ', '«', '»', '"', '“', '”');
                                    }
                                    else {
                                        var matchCompilerMarker = m_regexCompilerMarker.Match(text);
                                        compilerTestReady = matchCompilerMarker.Success;
                                    }
                                }
                            }
                            //сбор параграфов
                            if (collectPars) {
                                var stopCollecting = false;
                                if (collectSummary) {
                                    var m = m_regexNumberedHeader.Match(text);
                                    stopCollecting = m.Success && int.Parse(m.Groups[1].Value) >= 4 && //номер раздела как правило 5
                                                     (m.Groups[2].Value.Contains("учебно-методич", StringComparison.CurrentCultureIgnoreCase) || //в нем есть слово "методич"
                                                     par.MagicText.FirstOrDefault().formatting?.Bold == true); /* или он жирный */
                                    if (!stopCollecting) {
                                        if (//par.MagicText.FirstOrDefault().formatting?.Bold == true && 
                                            text.Contains("учебно-методич", StringComparison.CurrentCultureIgnoreCase) &&
                                            par.IsListItem) {
                                            stopCollecting = true;
                                        }
                                    }
                                }

                                if (stopCollecting) {
                                    if (collectSummary) rpd.SummaryParagraphs = parList;
                                    collectPars = false;
                                }
                                else {
                                    parList.Add(par);
                                }
                            }
                            //сбор текстовых значений
                            if (collectText) {
                                var stopCollecting = false;
                                var clearedText = text;
                                var matchNumberedItem = m_regexNumberedHeader.Match(text);
                                //if (collectQuestions) {
                                //    stopCollecting = matchNumberedItem.Success && int.Parse(matchNumberedItem.Groups[1].Value) >= 6 && //номер раздела как правило 8
                                //                    (matchNumberedItem.Groups[2].Value.Contains("Перечень", StringComparison.CurrentCultureIgnoreCase) || //в нем есть слово "Перечень"
                                //                    par.MagicText.FirstOrDefault().formatting?.Bold == true); /* или он жирный */
                                //}
                                if (!stopCollecting && collectBaseReferences) { //основная литература - конец списка на заголовке доп. литературы?
                                    stopCollecting = m_regexReferencesExtra.Any(r => r.IsMatch(text));
                                }
                                if (!stopCollecting && collectExtraReferences) { //доп. лит-ра - конец списка?
                                    stopCollecting = matchNumberedItem.Success && int.Parse(matchNumberedItem.Groups[1].Value) >= 7 && //номер раздела как правило 8
                                                     (matchNumberedItem.Groups[2].Value.Contains("Профессиональные", StringComparison.CurrentCultureIgnoreCase) || //в нем есть слово "Профессиональные"
                                                     par.MagicText.FirstOrDefault().formatting?.Bold == true); /* или он жирный */
                                }
                                if (!stopCollecting) {
                                    if (matchNumberedItem.Success) {
                                        clearedText = matchNumberedItem.Groups[2].Value.Trim(' ', '\r', '\n', '\t');
                                        var num = int.Parse(matchNumberedItem.Groups[1].Value);
                                        if (num < textList.Count) {
                                            stopCollecting = true;
                                        }
                                    }
                                    else { //не подошло по выражению - проверим как элемент нумерованного списка
                                        if (par.IsListItem && par.ListItemType == ListItemType.Numbered) {
                                            if (int.TryParse(par.GetListItemNumber(), out var num)) {
                                                if (num < textList.Count) {
                                                    stopCollecting = true;
                                                }
                                            }
                                        }
                                        //else {
                                        //    stopCollecting = true;
                                        //}
                                    }
                                }
                                if (stopCollecting) {
                                    if (collectQuestions) rpd.QuestionList = textList.ToList();
                                    if (collectBaseReferences) rpd.ReferencesBase = textList.ToList();
                                    if (collectExtraReferences) rpd.ReferencesExtra = textList.ToList();
                                    collectText = false;
                                    collectQuestions = false;
                                    collectBaseReferences = false;
                                    collectExtraReferences = false;
                                }
                                else {
                                    textList.Add(clearedText);
                                }
                                //}
                            }
                            //маркер начала содержания разделов и тем
                            if (rpd.SummaryParagraphs.Count == 0) {
                                if (m_regexSummaryBeginMarkers.Any(r => r.IsMatch(text))) {
                                    parList.Clear();
                                    collectPars = true;
                                    collectSummary = true;
                                }
                            }
                            //маркер вопросов к зкз/зачету
                            if (rpd.QuestionList.Count == 0) {
                                if (m_regexQuestionListBeginMarker.Any(r => r.IsMatch(text))) {
                                    textList.Clear();
                                    collectText = true;
                                    collectQuestions = true;
                                }
                            }
                            //маркер основной литературы
                            if (rpd.ReferencesBase.Count == 0) {
                                if (m_regexReferencesBase.Any(r => r.IsMatch(text))) {
                                    if (collectQuestions) {
                                        rpd.QuestionList = textList.ToList();
                                        collectQuestions = false;
                                    }
                                    textList.Clear();
                                    collectText = true;
                                    collectBaseReferences = true;
                                }
                            }
                            //маркер доп. литературы
                            if (rpd.ReferencesExtra.Count == 0) {
                                if (m_regexReferencesExtra.Any(r => r.IsMatch(text))) {
                                    textList.Clear();
                                    collectText = true;
                                    collectExtraReferences = true;
                                }
                            }
                            //направление подготовки
                            if (string.IsNullOrEmpty(rpd.DirectionCode)) {
                                var matchDirection = m_regexDirection.Match(text);
                                if (matchDirection.Success) {
                                    rpd.DirectionCode = string.Join("", matchDirection.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries));
                                    rpd.DirectionName = matchDirection.Groups[2].Value.Trim(' ', '«', '»', '"', '“', '”');
                                }
                            }
                            //формы обучения
                            //вариант 1: в двух абзацах
                            if (text.Equals("форма обучения", StringComparison.CurrentCultureIgnoreCase)) {
                                text += $" {par.NextParagraph.Text}";
                            }
                            //вариант 2: в одной строке
                            if (rpd.FormsOfStudy.Count == 0) {
                                var matchFormsOfStudy = m_regexFormsOfStudy.Match(text);
                                if (matchFormsOfStudy.Success) {
                                    var items = matchFormsOfStudy.Groups[1].Value.Split(',', StringSplitOptions.TrimEntries);
                                    foreach (var item in items) {
                                        switch (item.Trim().ToLower()) {
                                            case "очная":
                                                rpd.FormsOfStudy.Add(EFormOfStudy.FullTime);
                                                break;
                                            case "заочная":
                                                rpd.FormsOfStudy.Add(EFormOfStudy.PartTime);
                                                break;
                                            case "очно-заочная":
                                                rpd.FormsOfStudy.Add(EFormOfStudy.MixedTime);
                                                break;
                                            default:
                                                rpd.FormsOfStudy.Add(EFormOfStudy.Unknown);
                                                break;

                                        }
                                    }
                                }
                            }

                            //проверка на маркер [РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ]
                            if (string.IsNullOrEmpty(rpd.DisciplineName)) {
                                var matchRpdMarker = m_regexRpdMarker.Match(text);
                                disciplineTestReady = matchRpdMarker.Success;
                            }
                            //если встретился маркер конца титульной страницы
                            var matchPageEnd = m_regexTitlePageEndMarker.Match(text);
                            if (matchPageEnd.Success) {
                                rpd.Year = matchPageEnd.Groups[1].Value;
                            }
                            parIdx++;
                        }

                        var fullTimeTableIsOk = false;
                        var mixedTimeTableIsOk = false;
                        var partTimeTableIsOk = false;

                        //функция для нормализации текста: убираем лишние пробелы и приведение к UpperCase
                        Func<string, string> normalizeText = t => string.Join(" ", t.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)).ToUpper();

                        //проверка таблиц
                        foreach (var table in docx.Tables) {
                            var keepTestTable = true;
                            if (!rpd.EducationalWorks.Any()) {
                                keepTestTable = !App.TestForSummaryTableForEducationalWorks(table, rpd.EducationalWorks, PropertyAccess.Get);
                            }
                            if (keepTestTable && rpd.CompetenceMatrix == null) {
                                if (CompetenceMatrix.TestTable(table, out var format) && format == ECompetenceMatrixFormat.Rpd) {
                                    if (App.TestForTableOfCompetenceMatrix(table, format, out var matrix, out var errors)) {
                                        rpd.CompetenceMatrix = matrix;
                                        keepTestTable = false;
                                    }
                                    if (errors.Any()) rpd.Errors.AddRange(errors);
                                }
                                //testTable = !App.TestForTableOfCompetenceMatrix(table, rpd);
                            }
                            if (keepTestTable) {
                                EEvaluationTool[] evalTools = null;
                                string[][] studyResults = null;
                                if (App.TestForEduWorkTable(docx, table, rpd, PropertyAccess.Get, EEduWorkFixType.All, ref evalTools, ref studyResults,
                                                            0, null, null, /* параметры для режима исправления - здесь они не нужны */
                                                            out var formOfStudy)) {
                                    //получим список модулей обучения с оценочными средствами и компетенциями
                                    var eduWork = rpd.EducationalWorks[formOfStudy];

                                    eduWork.Modules = [];
                                    //var evalToolDic = Enum.GetValues(typeof(EEvaluationTool)).Cast<EEvaluationTool>().ToDictionary(x => x.GetDescription().ToUpper(), x => x);

                                    for (var row = eduWork.TableTopicStartRow; row <= eduWork.TableTopicLastRow; row++) { //минус ряд с "зачетом", минус ряд с "итого"
                                        var module = new StudyModule();
                                        var col = eduWork.TableColTopic; // startNumCol - 1;
                                        var cellIdx = col;
                                        while (col < eduWork.TableMaxColCount) {
                                            if (col == eduWork.TableColTopic) {
                                                if (table.Rows[row].Cells.Count <= cellIdx) {
                                                    var tt = 0;
                                                }
                                                try {
                                                    module.Topic = table.Rows[row].Cells[cellIdx].GetText();
                                                }
                                                catch (Exception exx) {
                                                    var tt = 0;
                                                }
                                            }
                                            if (col == eduWork.TableColEvalTools) { //оценочные средства
                                                module.EvaluationTools = [];
                                                //var counts = table.Rows.Select(r => r.Cells.Count).ToList();
                                                //module.Topic = table.Rows[row].Cells[col].GetText();
                                                foreach (var t in table.Rows[row].Cells[cellIdx].GetText().Split(',', '\n', ';')) {
                                                    //убираем лишние пробелы
                                                    if (Enums.EvalToolDic.TryGetValue(normalizeText(t), out var tool)) {
                                                        //var normalizedText = string.Join(" ", t.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)).ToUpper();
                                                        //if (Enum.TryParse(normalizedText, true, out EEvaluationTool tool)) {
                                                        module.EvaluationTools.Add(tool);
                                                    }
                                                    else {
                                                        if (string.IsNullOrEmpty(t)) {
                                                            rpd.Errors.Add(EErrorType.RpdMissingEvalTool, $"Тема [{module.Topic}] ({formOfStudy.GetDescription()})");
                                                        }
                                                        else {
                                                            rpd.Errors.Add(EErrorType.RpdUnknownEvalTool, $"Значение - {t}, тема [{module.Topic}] ({formOfStudy.GetDescription()})");
                                                        }
                                                    }
                                                }
                                            }
                                            if (col == eduWork.TableColCompetenceResults) { //результаты обучения - компетенции
                                                module.CompetenceResultCodes = table.Rows[row].Cells[cellIdx].GetText(",").Split(',', '\n', ';').ToHashSet();
                                            }
                                            col += table.Rows[row].Cells[cellIdx].GridSpan + 1; //переход на след. ячейку с учетом объединений
                                            cellIdx++;                                          //индексы ячеек последовательны
                                        }
                                        eduWork.Modules.Add(module);
                                    }

                                    keepTestTable = false;  //таблица обработана
                                }
                            }

                        }
                        //итоговая проверка
                        if (rpd.FormsOfStudy.Count != rpd.EducationalWorks.Count) {
                            foreach (var form in rpd.FormsOfStudy) {
                                if (!rpd.EducationalWorks.ContainsKey(form)) {
                                    rpd.Errors.Add(EErrorType.RpdMissingEduWork, $"Форма обучения [{form}]");
                                }
                            }
                        }
                        if (rpd.CompetenceMatrix == null || !rpd.CompetenceMatrix.IsLoaded) {
                            rpd.Errors.Add(EErrorType.RpdMissingCompetenceMatrix);
                        }
                        if (rpd.SummaryParagraphs.Count == 0) {
                            rpd.Errors.Add(EErrorType.RpdMissingTOC);
                        }
                        if (rpd.QuestionList.Count == 0) {
                            rpd.Errors.Add(EErrorType.RpdMissingQuestionList);
                        }
                        if (rpd.EducationalWorks.Count != rpd.FormsOfStudy.Count) {
                            rpd.Errors.Add(EErrorType.RpdNotFullEduWorks, $"(только для: {string.Join(", ", rpd.EducationalWorks?.Select(x => x.Key.GetDescription()))})");
                        }
                        if (rpd.ReferencesBase.Count == 0) {
                            rpd.Errors.Add(EErrorType.RpdMissingReferencesBase);
                        }
                        if (rpd.ReferencesExtra.Count == 0) {
                            rpd.Errors.Add(EErrorType.RpdMissingReferencesExtra);
                        }
                        foreach (var eduWork in rpd.EducationalWorks) {
                            if (eduWork.Value.Table == null) {
                                rpd.Errors.Add(EErrorType.RpdMissingEduWorkTable, $"(форма обучения [{eduWork.Key.GetDescription()}])");
                            }
                        }
                        if (string.IsNullOrEmpty(rpd.Department)) rpd.Errors.Add(EErrorType.RpdMissingDepartment);
                        if (string.IsNullOrEmpty(rpd.Profile)) rpd.Errors.Add(EErrorType.RpdMissingProfile);
                        if (string.IsNullOrEmpty(rpd.Year)) rpd.Errors.Add(EErrorType.RpdMissingYear);
                        if (string.IsNullOrEmpty(rpd.DirectionCode)) rpd.Errors.Add(EErrorType.RpdMissingDirectionCode);
                        if (string.IsNullOrEmpty(rpd.DirectionName)) rpd.Errors.Add(EErrorType.RpdMissingDirectionName);
                        if (string.IsNullOrEmpty(rpd.DisciplineName)) rpd.Errors.Add(EErrorType.RpdMissingDisciplineName);
                    }
                }
            }
            catch (Exception ex) {
                rpd.Errors.AddException(ex);
            }

            return rpd;
        }

        /// <summary>
        /// Установить новый список предыд. дисциплин
        /// </summary>
        /// <param name="prevNames"></param>
        /// <exception cref="NotImplementedException"></exception>
        internal void SetPrevDisciplines(IEnumerable<string> prevNames) {
            if (prevNames?.Any() ?? false) {
                PrevDisciplines = string.Join(", ", prevNames.Select(n => $"«{n}»"));
            }
        }

        /// <summary>
        /// Установить новый список последующих дисциплин
        /// </summary>
        /// <param name="nextNames"></param>
        /// <exception cref="NotImplementedException"></exception>
        internal void SetNextDisciplines(IEnumerable<string> nextNames) {
            if (nextNames?.Any() ?? false) {
                NextDisciplines = string.Join(", ", nextNames.Select(n => $"«{n}»"));
            }
        }

        /// <summary>
        /// Получить список имён предыд. дисциплин
        /// </summary>
        /// <returns></returns>
        internal HashSet<string> GetPrevDisciplines() {
            var names = new HashSet<string>();

            if (!string.IsNullOrEmpty(PrevDisciplines) && string.Compare(PrevDisciplines,"{PrevDisciplines}") != 0) {
                names = PrevDisciplines.Split(",", StringSplitOptions.TrimEntries).Select(d => d.Trim('«', '»')).ToHashSet();
            }

            return names;
        }

        /// <summary>
        /// Получить список имён предыд. дисциплин
        /// </summary>
        /// <returns></returns>
        internal HashSet<string> GetNextDisciplines() {
            var names = new HashSet<string>();

            if (!string.IsNullOrEmpty(NextDisciplines) && string.Compare(NextDisciplines, "{NextDisciplines}") != 0) {
                names = NextDisciplines.Split(",", StringSplitOptions.TrimEntries).Select(d => d.Trim('«', '»')).ToHashSet();
            }

            return names;
        }
    }
}
