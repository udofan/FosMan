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
    internal class Rpd {
        //Кафедра
        static Regex m_regexDepartment = new(@"Кафедра\s+(.+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //Маркер конца титульной странцы "Москва 20xx"
        static Regex m_regexTitlePageEndMarker = new(@"Москва\s+(\d{4})", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //профиль/направленность подготовки
        static Regex m_regexProfileInline = new(@"(Профиль|Направленность\s+подготовки)[:]*\s+[«""“]+(.+)[»""”]+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexProfileMarker = new(@"(Профиль|Направленность\s+подготовки)[:]*\s*[«""“]*(.*)[»""”]*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //наименование в кавычках
        static Regex m_regexNameQuoted = new(@"[«""“]+(.+)[»""”]+$", RegexOptions.Compiled);
        //наименование в опциональных кавычках
        static Regex m_regexName = new(@"[«""“]*(.+)[»""”]*$", RegexOptions.Compiled);
        //направление подготовки
        static Regex m_regexDirection = new(@"(\d{2}\s*\.\s*\d{2}\s*\.\s*\d{2})\s+(.*)$", RegexOptions.Compiled);
        //форма обучения
        static Regex m_regexFormsOfStudy = new(@"Форма\s+обучения[:]*\s+(.+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //маркер [РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ]
        static Regex m_regexRpdMarker = new(@"рабочая\s+программа\s+дисциплины", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Факультет
        /// </summary>
        public string Faculty { get; set; }
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
                                    var matchProfileMarket = m_regexProfileMarker.Match(text);
                                    if (matchProfileMarket.Success) {
                                        if (profileField.Length > 0) profileField += " ";
                                        profileField += matchProfileMarket.Groups[2].Value;
                                        profileTestReady = true;
                                    }
                                    profileTestReady = matchProfileMarket.Success;
                                }
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

                    //проверка таблиц
                    foreach (var table in docx.Tables) {
                        if (!rpd.EducationalWorks.Any()) {
                            App.TestTableForEducationalWorks(table, rpd.EducationalWorks, true);
                        }
                        if (rpd.CompetenceMatrix == null) {
                            TestTableForCompetenceMatrix(table, rpd);
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
        private static void TestTableForCompetenceMatrix(Table table, Rpd rpd) {
            var matrix = new CompetenceMatrix() {
                Items = [],
                Errors = [],
            };
            if (CompetenceMatrix.TryParseTable(table, matrix)) {
                rpd.CompetenceMatrix = matrix;
                rpd.Errors.AddRange(matrix.Errors);
            }
        }
    }
}
