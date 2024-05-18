using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.WebSockets;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace FosMan {
    internal class Rpd : BaseObj{
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
        };
        static List<(Regex, int)> m_regexNextDisciplines = new() {
            //последующих учебных дисциплин
            //изучении последующих профессиональных дисциплин
            (new(@"послед[^.]+(учеб)?[^.]+дисциплин[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 2),
            //изучения следующих дисциплин
            (new(@"след[^.]+(учеб)?[^.]+дисциплин[:]*\s+([^.]+).", RegexOptions.IgnoreCase | RegexOptions.Compiled), 2),
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
        static public Regex RegexFullTimeTable = new(@"^очная\s+форма\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        static public Regex RegexMixedTimeTable = new(@"^очно-заочная\s+форма\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        static public Regex RegexPartTimeTable = new(@"^заочная\s+форма\s+", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        /// <summary>
        /// Факультет
        /// </summary>
        //public string Faculty { get; set; }
        /// <summary>
        /// Год
        /// </summary>
        public string Year { get; set; }
        /// <summary>
        /// Дисциплина
        /// </summary>
        public string DisciplineName { get; set; }
        /// <summary>
        /// Профиль
        /// </summary>
        public string Profile { get; set; }
        /// <summary>
        /// Кафедра
        /// </summary>
        public string Department { get; set; }
        /// <summary>
        /// Код направления подготовки
        /// </summary>
        public string DirectionCode { get; set; }
        /// <summary>
        /// Направление подготовки
        /// </summary>
        public string DirectionName { get; set; }
        /// <summary>
        /// Формы обучения
        /// </summary>
        public List<EFormOfStudy> FormsOfStudy { get; set; }
        /// <summary>
        /// Матрица компетенций (п. 1 в РПД)
        /// </summary>
        public CompetenceMatrix CompetenceMatrix { get; set; }
        /// <summary>
        /// Исходный docx-файл
        /// </summary>
        public string SourceFileName { get; set; }
        /// <summary>
        /// Выявленные ошибки
        /// </summary>
        public List<string> Errors { get; set; }
        /// <summary>
        /// Учебная работа по формам обучения
        /// </summary>
        public Dictionary<EFormOfStudy, EducationalWork> EducationalWorks { get; set; }
        /// <summary>
        /// Список доп. ошибок, выявленных при проверке
        /// </summary>
        public List<string> ExtraErrors { get; set; }
        /// <summary>
        /// Составитель
        /// </summary>
        public string Compiler { get; set; }
        /// <summary>
        /// Предшествующие дисциплины
        /// </summary>
        public string PrevDisciplines { get; set; }
        /// <summary>
        /// Последующие дисциплины
        /// </summary>
        public string NextDisciplines { get; set; }

        /// <summary>
        /// Загрузка РПД из файла
        /// </summary>
        /// <param name="fileName"></param>
        public static Rpd LoadFromFile(string fileName) {
            var rpd = new Rpd() {
                Errors = [], 
                SourceFileName = fileName, 
                EducationalWorks = [], 
                FormsOfStudy = []
            };

            try {
                using (var docx = DocX.Load(fileName)) {
                    //парсинг титульной страницы
                    var parIdx = 0;
                    var disciplineTestReady = false;    //флаг готовности поиска имени дисциплины
                    var disciplineName = "";
                    var profileTestReady = false;
                    var profileField = "";              //если профиль разнесен по неск. строкам, здесь он будет накапливаться
                    var compilerTestReady = false;
                    var fullTimeTestTableReady = false;
                    var mixedTimeTestTableReady = false;
                    var partTimeTestTableReady = false;

                    foreach (var par in docx.Paragraphs) {
                        var text = par.Text.Trim();
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
                        //таблица очная форма обучения
                        //if (rpd.EducationalWorks.TryGetValue(EFormOfStudy.FullTime, out var eduWork)) {
                        //    if (eduWork.Table == null) {
                        //        if (fullTimeTestTableReady) {
                        //            if ((par.FollowingTables?.Count ?? 0) == 1) {
                        //                eduWork.Table = par.FollowingTables[0];
                        //                fullTimeTestTableReady = false;
                        //            }
                        //        }
                        //        if (m_regexFullTimeTable.IsMatch(text)) {
                        //            fullTimeTestTableReady = true;
                        //        }
                        //    }
                        //}
                        ////таблица очно-заочная форма обучения
                        //if (rpd.EducationalWorks.TryGetValue(EFormOfStudy.MixedTime, out eduWork)) {
                        //    if (eduWork.Table == null) {
                        //        if (mixedTimeTestTableReady) {
                        //            if ((par.FollowingTables?.Count ?? 0) == 1) {
                        //                eduWork.Table = par.FollowingTables[0];
                        //                mixedTimeTestTableReady = false;
                        //            }
                        //        }
                        //        if (m_regexMixedTimeTable.IsMatch(text)) {
                        //            mixedTimeTestTableReady = true;
                        //        }
                        //    }
                        //}
                        ////таблица заочная форма обучения
                        //if (rpd.EducationalWorks.TryGetValue(EFormOfStudy.PartTime, out eduWork)) {
                        //    if (eduWork.Table == null) {
                        //        if (partTimeTestTableReady) {
                        //            if ((par.FollowingTables?.Count ?? 0) == 1) {
                        //                eduWork.Table = par.FollowingTables[0];
                        //                partTimeTestTableReady = false;
                        //            }
                        //        }
                        //        if (m_regexPartTimeTable.IsMatch(text)) {
                        //            partTimeTestTableReady = true;
                        //        }
                        //    }
                        //}
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

                    //проверка таблиц
                    foreach (var table in docx.Tables) {
                        var testTable = true;
                        if (!rpd.EducationalWorks.Any()) {
                            testTable = !App.TestTableForEducationalWorks(table, rpd.EducationalWorks, true);
                        }
                        if (testTable && rpd.CompetenceMatrix == null) {
                            testTable = !TestTableForCompetenceMatrix(table, rpd);
                        }
                        if (testTable) {
                            //проверка на таблицы учебных работ с темами
                            var par = table.Paragraphs.FirstOrDefault();
                            for (var i = 0; i < 3; i++) {
                                par = par.PreviousParagraph;
                                if (RegexFullTimeTable.IsMatch(par.Text)) {
                                    rpd.EducationalWorks[EFormOfStudy.FullTime].Table = table;
                                    break;
                                }
                                if (RegexMixedTimeTable.IsMatch(par.Text)) {
                                    rpd.EducationalWorks[EFormOfStudy.MixedTime].Table = table; 
                                    break;
                                }
                                if (RegexPartTimeTable.IsMatch(par.Text)) {
                                    rpd.EducationalWorks[EFormOfStudy.PartTime].Table = table;
                                    break;
                                }
                            }
                        }

                    }
                    //итоговая проверка
                    if (rpd.FormsOfStudy.Count != rpd.EducationalWorks.Count) {
                        foreach (var form in rpd.FormsOfStudy) {
                            if (!rpd.EducationalWorks.ContainsKey(form)) {
                                rpd.Errors.Add($"Не найдена информация по учебным работам по форме обучения {form}");
                            }
                        }
                    }
                    if (rpd.CompetenceMatrix == null || !rpd.CompetenceMatrix.IsLoaded) {
                        rpd.Errors.Add("Не найдена матрица компетенций.");
                    }
                    if (string.IsNullOrEmpty(rpd.Department)) rpd.Errors.Add("Не удалось определить название кафедры");
                    if (string.IsNullOrEmpty(rpd.Profile)) rpd.Errors.Add("Не удалось определить профиль");
                    if (string.IsNullOrEmpty(rpd.Year)) rpd.Errors.Add("Не удалось год программы");
                    if (string.IsNullOrEmpty(rpd.DirectionCode)) rpd.Errors.Add("Не удалось шифр направления подготовки");
                    if (string.IsNullOrEmpty(rpd.DirectionName)) rpd.Errors.Add("Не удалось наименование направления подготовки");
                    if (string.IsNullOrEmpty(rpd.DisciplineName)) rpd.Errors.Add("Не удалось название дисциплины");
                }
            }
            catch (Exception ex) {
                rpd.Errors.Add($"{ex.Message}\r\n{ex.StackTrace}");
            }

            return rpd;
        }

        /// <summary>
        /// Проверка таблицы на матрицу компетенций
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rpd"></param>
        private static bool TestTableForCompetenceMatrix(Table table, Rpd rpd) {
            var matrix = new CompetenceMatrix() {
                Items = [],
                Errors = [],
            };
            var result = CompetenceMatrix.TryParseTable(table, matrix);
            if (result) {
                rpd.CompetenceMatrix = matrix;
                rpd.Errors.AddRange(matrix.Errors);
            }

            return result;
        }
    }
}
