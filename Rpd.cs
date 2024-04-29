using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.WebSockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace FosMan {
    internal class Rpd {
        //Кафедра
        static Regex m_regexDepartment = new(@"(Кафедра.+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //Маркер конца титульной странцы "Москва 20xx"
        static Regex m_regexTitlePageEndMarker = new(@"Москва\s+\d{4}", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //профиль
        static Regex m_regexProfile = new(@"Профиль\s+[«""“](.+)[»""”]", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //дисциплина
        static Regex m_regexDiscipline = new(@"[«""“](.+)[»""”]", RegexOptions.Compiled);
        //направление подготовки
        static Regex m_regexDirection = new(@"(\d{2}\s*\.\s*\d{2}\s*\.\s*\d{2})\s+(.*)$", RegexOptions.Compiled);
        //форма обучения
        static Regex m_regexFormsOfStudy = new(@"Форма\s+обучения:\s+(.+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //маркер [РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ]
        static Regex m_regexRpdMarker = new(@"рабочая\s+программа\s+дисциплины", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Факультет
        /// </summary>
        public string Faculty { get; set; }
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
        /// Загрузка РПД из файла
        /// </summary>
        /// <param name="fileName"></param>
        public static Rpd LoadFromFile(string fileName) {
            var rpd = new Rpd() {
                Errors = [], 
                SourceFileName = fileName, 
                FormsOfStudy = []
            };

            try {
                using (var docx = DocX.Load(fileName)) {
                    //парсинг титульной страницы
                    var parIdx = 0;
                    var disciplineTestReady = false;    //флаг готовности поиска имени дисциплины
                    foreach (var par in docx.Paragraphs) {
                        //кафедра
                        var matchDepartment = m_regexDepartment.Match(par.Text);
                        if (matchDepartment.Success) rpd.Department = matchDepartment.Groups[1].Value.Trim();
                        //профиль
                        var matchProfile = m_regexProfile.Match(par.Text);
                        if (matchProfile.Success) rpd.Profile = matchProfile.Groups[1].Value.Trim();
                        //направление подготовки
                        var matchDirection = m_regexDirection.Match(par.Text);
                        if (matchDirection.Success) {
                            rpd.DirectionCode = string.Join("", matchDirection.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries));
                            rpd.DirectionName = matchDirection.Groups[2].Value.Trim(' ', '«', '»', '"', '“', '”');
                        }
                        //формы обучения
                        var matchFormsOfStudy = m_regexFormsOfStudy.Match(par.Text);
                        if (matchFormsOfStudy.Success) {
                            var items = matchFormsOfStudy.Groups[1].Value.Split(',');
                            foreach (var item in items) {
                                switch (item.Trim().ToLower()) {
                                    case "очная":
                                        rpd.FormsOfStudy.Add(EFormOfStudy.FullTime);
                                        break;
                                    case "заочная":
                                        rpd.FormsOfStudy.Add(EFormOfStudy.PartTime);
                                        break;
                                    case "очно-заочная":
                                        rpd.FormsOfStudy.Add(EFormOfStudy.FullTimeAndPartTime);
                                        break;
                                    default:
                                        rpd.FormsOfStudy.Add(EFormOfStudy.Unknown);
                                        break;

                                }
                            }
                        }

                        if (disciplineTestReady) {
                            var matchDiscipline = m_regexDiscipline.Match(par.Text);
                            if (matchDiscipline.Success) {
                                rpd.DisciplineName = matchDiscipline.Groups[1].Value.Trim();
                            }
                        }

                        //проверка на маркер [РАБОЧАЯ ПРОГРАММА ДИСЦИПЛИНЫ]
                        var matchRpdMarker = m_regexRpdMarker.Match(par.Text);
                        disciplineTestReady = matchRpdMarker.Success;

                        //если встретился маркер конца титульной страницы
                        if (m_regexTitlePageEndMarker.IsMatch(par.Text)) {
                            break;
                        }
                        parIdx++;
                    }
                }
            }
            catch (Exception ex) {
                rpd.Errors.Add($"{ex.Message}\r\n{ex.StackTrace}");
            }

            return rpd;
        }
    }
}
