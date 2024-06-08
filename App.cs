using FastMember;
using Microsoft.Web.WebView2.Core;
using System;
using System.CodeDom;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.Eventing.Reader;
using System.DirectoryServices.ActiveDirectory;
using System.Drawing;
using System.Drawing.Design;
using System.Globalization;
using System.Linq;
using System.Numerics;
using System.Reflection;
using System.Security.Policy;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Text.Unicode;
using System.Threading.Tasks;
using System.Web;
using Xceed.Document.NET;
using Xceed.Words.NET;
using static System.Resources.ResXFileRef;

namespace FosMan {
    /// <summary>
    /// Тип доступа к свойству
    /// </summary>
    public enum PropertyAccess {
        Get,    //получить значение
        Set     //установить значение
    }
    /// <summary>
    /// Загруженные данные
    /// </summary>
    static internal class App {
        //const string FIXED_RPD_DIRECTORY = "FixedRpd";
        //const string CONFIG_FILENAME = "appconfig.json";

        //таблица учебных работ по формам обучения
        static Regex m_regexEduWorkType = new(@"вид.+ учеб.+ работ", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexFormFullTime = new(@"^очная", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexFormMixedTime = new(@"^очно-заоч", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexFormPartTime = new(@"^заочная", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexTotalHours = new(@"(общ|всего)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexContactHours = new(@"(аудит|контакт)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexLectureHours = new(@"лекц", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexLabHours = new(@"лабор", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexPracticalHours = new(@"практич", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexSelfStudyHours = new(@"самост", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexControlHours = new(@"контроль", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexControlForm = new(@"форма", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //тест на вид итогового контроля
        static Regex m_regexControlExam = new(@"экзам", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexControlTest = new(@"зач", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexControlTestGrade = new(@"зач.+с.+оц", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        //static Regex m_regexTestTableCaptionFullTimeEduWorks;
        //static Regex m_regexTestTableEduWorks = new(@"Наименован.*раздел.*тем", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        //для выявления раздела
        static Regex m_regexRpdChapter = new(@"^(\d+).\s+(.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        //выражения для проверка заголовков таблиц учебных работ (содержания) по формам обучения
        static Dictionary<EFormOfStudy, Regex> m_eduWorkTableHeaders = new() {
            [EFormOfStudy.FullTime] = new(@"^очная\s+форма($|\s+)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            [EFormOfStudy.MixedTime] = new(@"^очно-заочная\s+форма($|\s+)", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            [EFormOfStudy.PartTime] = new(@"^заочная\s+форма($|\s+)", RegexOptions.IgnoreCase | RegexOptions.Compiled)
        };

        static CompetenceMatrix m_competenceMatrix = null;
        static Dictionary<string, Curriculum> m_curriculumDic = [];
        static ConcurrentDictionary<string, Rpd> m_rpdDic = [];
        static Dictionary<string, CurriculumGroup> m_curriculumGroupDic = [];
        //static Dictionary<string, Department> m_departments = [];
        static Config m_config = new();
        static JsonSerializerOptions m_jsonOptions = new() {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic),
            Converters = {
                new JsonStringEnumConverter()
            }
        };

        public delegate void RpdAddEventHandler(Rpd rpd);
        public static event RpdAddEventHandler RpdAdd = null;

        public static CompetenceMatrix CompetenceMatrix { get => m_competenceMatrix; }

        public static ConcurrentDictionary<string, Rpd> RpdList { get => m_rpdDic; }

        /// <summary>
        /// Конфигурация
        /// </summary>
        public static Config Config { get => m_config; }

        /// <summary>
        /// Учебные планы в сторе
        /// </summary>
        public static Dictionary<string, Curriculum> Curricula { get => m_curriculumDic; }
        /// <summary>
        /// Группы УП
        /// </summary>
        public static Dictionary<string, CurriculumGroup> CurriculumGroups { get => m_curriculumGroupDic; }

        /// <summary>
        /// Список кафедр
        /// </summary>
        //public static Dictionary<string, Department> Departments { get => m_departments; }

        public static bool HasCurriculumFile(string fileName) => m_curriculumDic.ContainsKey(fileName);

        public static bool HasRpdFile(string fileName) => m_rpdDic.ContainsKey(fileName);

        /// <summary>
        /// Добавить УП в общий стор
        /// </summary>
        /// <param name="curriculum"></param>
        /// <returns></returns>
        static public bool AddCurriculum(Curriculum curriculum) {
            var result = false;
            if (m_curriculumDic.TryAdd(curriculum.SourceFileName, curriculum)) {
                if (!m_curriculumGroupDic.TryGetValue(curriculum.GroupKey, out var group)) {
                    group = new() {
                        Department = curriculum.Department,
                        DirectionCode = curriculum.DirectionCode,
                        DirectionName = curriculum.DirectionName,
                        Profile = curriculum.Profile,
                        FSES = curriculum.FSES
                    };
                    m_curriculumGroupDic[curriculum.GroupKey] = group;
                }
                group.AddCurriculum(curriculum);
                CheckDisciplines();
                result = true;
            }
            return result;
        }

        static public bool AddRpd(Rpd rpd) {
            var result = m_rpdDic.TryAdd(rpd.SourceFileName, rpd);
            if (result) {
                RpdAdd?.Invoke(rpd);
            }
            return result;
        }

        /// <summary>
        /// Проверка дисциплин на списки компетенций
        /// </summary>
        static void CheckDisciplines() {
            foreach (var curr in m_curriculumDic.Values) {
                foreach (var disc in curr.Disciplines.Values) {
                    disc.CheckCompetences(m_competenceMatrix);
                }
            }
        }

        static public void SetCompetenceMatrix(CompetenceMatrix matrix) {
            m_competenceMatrix = matrix;
            CheckDisciplines();
        }

        /// <summary>
        /// Extension для ячейки: получение текста по всем абзацам, объединенными заданной строкой
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="joinParText"></param>
        /// <returns></returns>
        static public string GetText(this Cell cell, string joinParText = " ", bool applyTrim = true) {
            var text = string.Join(joinParText, cell.Paragraphs.Select(p => p.Text));
            if (applyTrim) {
                return text.Trim();
            }
            return text;
        }

        static void AddError(this StringBuilder report, string error) {
            report.Append($"<div style='color: red'>{error}</div>");
        }

        static void AddDiv(this StringBuilder report, string Summary) {
            report.Append($"<div>{Summary}</div>");
        }

        static void AddTocElement(this StringBuilder toc, string element, string anchor, int errorCount) {
            var color = errorCount > 0 ? "red" : "green";
            toc.Append($"<li><a href='#{anchor}'><span style='color:{color};'>{element} (ошибок: {errorCount})</span></a></li>");
        }

        static void AddDisciplineCheckTableRow(this StringBuilder table, int rowNum, string description, bool result, string comment) {
            var tdStyle = " style='border: 1px solid;'";
            var style = $" style='color:{(result ? "green" : "red")}'";
            var resultMark = result ? "&check;" : "&times;";
            table.Append($"<tr {style}>");
            table.Append($"<td {tdStyle}>{rowNum}</td><td {tdStyle}>{description}</td><td {tdStyle}>{resultMark}</td><td {tdStyle}>{comment}</td>");
            table.Append("</tr>");
        }

        /// <summary>
        /// Проверка дисциплины по учебному плану
        /// </summary>
        /// <param name="rpd"></param>
        /// <param name="discipline"></param>
        /// <param name="repTable"></param>
        static bool ApplyDisciplineCheck(EFormOfStudy formOfStudy, Rpd rpd, CurriculumDiscipline discipline, StringBuilder repTable,
                                         ref int pos, ref int errorCount, string description, Func<EducationalWork, (bool result, string msg)> func) {
            pos++;

            rpd.EducationalWorks.TryGetValue(formOfStudy, out EducationalWork eduWork);
            var ret = func.Invoke(eduWork);
            if (!ret.result) errorCount++;

            repTable.AddDisciplineCheckTableRow(pos, description, ret.result, ret.msg);

            return ret.result;
        }

        /// <summary>
        /// Проверить загруженные РПД
        /// </summary>
        public static void CheckRdp(List<Rpd> rpdList, out string htmlReport) {
            htmlReport = string.Empty;

            StringBuilder html = new("<html><body><h2>Отчёт по проверке РПД</h2>");
            StringBuilder toc = new("<div><ul>");
            StringBuilder rep = new("<div>");
            var idx = 0;
            foreach (var rpd in rpdList) {
                var anchor = $"rpd{idx}";
                idx++;
                rpd.ExtraErrors = [];
                var errorCount = 0;

                //var discipName = rpd.DisciplineName;
                //if (string.IsNullOrEmpty(discipName)) discipName = "?";
                rep.Append($"<div id='{anchor}' style='width: 100%;'><h3 style='background-color: lightsteelblue'>{rpd.DisciplineName ?? "?"}</h3>");
                rep.Append("<div style='padding-left: 30px'>");
                rep.AddDiv($"Файл РПД: <b>{rpd.SourceFileName}</b>");
                //ошибки, выявленные при загрузке
                if (rpd.Errors.Any()) {
                    errorCount += rpd.Errors.Count;
                    rep.Append("<p />");
                    rep.AddDiv($"<div style='color: red'>Ошибки, выявленные при загрузке РПД ({rpd.Errors.Count} шт.):</div>");
                    rpd.Errors.ForEach(e => rep.AddError(e));
                }
                //ищем Учебные планы
                var curricula = FindCurricula(rpd);
                if (curricula != null && curricula.Any()) {
                    //проверка на отсутствие УПов
                    var missedFormsOfStudy = rpd.FormsOfStudy.ToHashSet();
                    missedFormsOfStudy.RemoveWhere(x => curricula.ContainsKey(x));
                    foreach (var item in missedFormsOfStudy) {
                        rep.Append("<p />");
                        rep.AddDiv($"Форма обучения: <b>{item.GetDescription()}</b>");
                        rep.AddDiv($"Файл УП: <b><span style='color: red'>Не найден УП для данной формы обучения</span></b>");
                        rep.AddError("Проверка РПД по УП невозможна.");
                        //rep.AddError($"Не найден УП для формы обучения <b>{item.GetDescription()}</b>.");
                        errorCount++;
                    }

                    //поиск дисциплины в УПах
                    foreach (var curriculum in curricula) {
                        rep.Append("<p />");
                        rep.AddDiv($"Форма обучения: <b>{curriculum.Value.FormOfStudy.GetDescription()}</b>");
                        rep.AddDiv($"Файл УП: <b>{curriculum.Value.SourceFileName}</b>");
                        var discipline = curriculum.Value.FindDiscipline(rpd.DisciplineName);
                        var checkPos = 0;
                        if (discipline != null) {
                            var tdStyle = " style='border: 1px solid;'";
                            var table = new StringBuilder(@"<table style='border: 1px solid'><tr style='font-weight: bold; background-color: lightgray'>");
                            table.Append($"<th {tdStyle}>№ п/п</th><th {tdStyle}>Проверка</th><th {tdStyle}>Результат</th><th {tdStyle}>Комментарий</th>");
                            table.Append("</tr>");
                            if (ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Наличие данных для формы обучения", (eduWork) => {
                                var result = eduWork != null;
                                var msg = result ? "" : "В описании РПД не определена учебная работа для данной формы обучения";
                                return (result, msg);
                            })) {
                                //проверка времен учебной работы
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Итоговое время", (eduWork) => {
                                    var result = discipline.EducationalWork.TotalHours == eduWork.TotalHours;
                                    var msg = result ? "" : $"Итоговое время [{discipline.EducationalWork.TotalHours}] не соответствует УП (д.б. {eduWork.TotalHours}).";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Время контактной работы", (eduWork) => {
                                    var result = (discipline.EducationalWork.ContactWorkHours ?? 0) == (eduWork.ContactWorkHours ?? 0);
                                    var msg = result ? "" : $"Время контактной работы [{discipline.EducationalWork.ContactWorkHours}] не соответствует УП (д.б. {eduWork.ContactWorkHours}).";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Время контроля", (eduWork) => {
                                    var result = (discipline.EducationalWork.ControlHours ?? 0) == (eduWork.ControlHours ?? 0);
                                    var msg = result ? "" : $"Время контроля [{discipline.EducationalWork.ControlHours}] не соответствует УП (д.б. {eduWork.ControlHours}).";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Время самостоятельной работы", (eduWork) => {
                                    var result = (discipline.EducationalWork.SelfStudyHours ?? 0) == (eduWork.SelfStudyHours ?? 0);
                                    var msg = result ? "" : $"Время самостоятельных работ [{discipline.EducationalWork.SelfStudyHours}] не соответствует УП (д.б. {eduWork.SelfStudyHours}).";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Время практических работ", (eduWork) => {
                                    var result = (discipline.EducationalWork.PracticalHours ?? 0) == (eduWork.PracticalHours ?? 0);
                                    var msg = result ? "" : $"Время практических работ [{discipline.EducationalWork.PracticalHours}] не соответствует УП (д.б. {eduWork.PracticalHours}).";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Время лекций", (eduWork) => {
                                    var result = (discipline.EducationalWork.LectureHours ?? 0) == (eduWork.LectureHours ?? 0);
                                    var msg = result ? "" : $"Время лекций [{discipline.EducationalWork.LectureHours}] не соответствует УП (д.б. {eduWork.LectureHours}).";
                                    return (result, msg);
                                });
                                //проверка итогового контроля
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Итоговый контроль", (eduWork) => {
                                    var error = false;
                                    var msg = "";
                                    if (discipline.EducationalWork.ControlForm == EControlForm.Unknown) {
                                        msg = "Не определен тип контроля.";
                                    }
                                    else if (eduWork.ControlForm == EControlForm.Unknown) {
                                        msg = $"В УП не определен тип контроля."; // [{eduWork.ControlForm.GetDescription()}] в УП не определено значение.";
                                    }
                                    else if (discipline.EducationalWork.ControlForm != eduWork.ControlForm) {
                                        msg = $"Тип контроля [{discipline.EducationalWork.ControlForm.GetDescription()}] не соответствует " +
                                              $"типу контроля из УП - [{eduWork.ControlForm.GetDescription()}]";
                                    }
                                    return (!error, msg);
                                });
                            }

                            //проверка компетенций
                            var checkCompetences = true;
                            checkCompetences &= ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Наличие матрицы компетенций в УП", (eduWork) => {
                                var result = discipline.CompetenceList?.Any() ?? false;
                                var msg = result ? "" : $"Не удалось определить матрицу компетенций в УП.";
                                return (result, msg);
                            });
                            checkCompetences &= ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Проверка матрицы компетенций в РПД", (eduWork) => {
                                var result = rpd.CompetenceMatrix?.IsLoaded ?? false;
                                var msg = result ? "" : $"В РПД не обнаружена матрица компетенций.";
                                if (result && rpd.CompetenceMatrix.Errors.Any()) {
                                    msg = "В матрице обнаружены ошибки:<br />";
                                    msg += string.Join("<br />", rpd.CompetenceMatrix.Errors);
                                    result = false;
                                }
                                if (!result) {
                                    if (msg.Length > 0) msg += "<br />";
                                    msg += "Проверка компетенций невозможна";
                                }
                                return (result, msg);
                            });
                            /*
                            checkCompetences &= ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Проверка загруженной матрицы компетенций в программу", (eduWork) => {
                                var result = m_competenceMatrix?.IsLoaded ?? false;
                                var msg = result ? "" : $"Матрица компетенций не загружена в программу.";
                                if (result && m_competenceMatrix.Errors.Any()) {
                                    msg = "В матрице обнаружены ошибки:<br />";
                                    msg += string.Join("<br />", m_competenceMatrix.Errors);
                                    result = false;
                                }
                                if (!result) {
                                    if (msg.Length > 0) msg += "<br />";
                                    msg += "Проверка компетенций невозможна";
                                }
                                return (result, msg);
                            });
                            */

                            if (checkCompetences) {
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Проверка матрицы компетенций", (eduWork) => {
                                    var achiCodeList = rpd.CompetenceMatrix.GetAllAchievementCodes();
                                    var Summary = "";
                                    var matrixError = false;
                                    foreach (var code in discipline.CompetenceList) {
                                        var elem = "";
                                        if (achiCodeList.Contains(code)) {
                                            elem = $"<span style='color: green; font-weight: bold'>{code}</span>";
                                        }
                                        else {
                                            elem = $"<span style='color: red; font-weight: bold'>{code}</span>";
                                            matrixError = true;
                                        }

                                        if (Summary.Length > 0) Summary += "; ";
                                        Summary += elem;
                                    }
                                    foreach (var missedCode in achiCodeList.Except(discipline.CompetenceList)) {
                                        matrixError = true;
                                        if (Summary.Length > 0) Summary += "; ";
                                        Summary += $"<span style='color: red; font-decoration: italic'>{missedCode}??</span>"; ;
                                    }
                                    var msg = matrixError ? $"Выявлено несоответствие компетенций: {Summary}" : "";
                                    return (!matrixError, msg);
                                });
                            }
                            table.Append("</table>");
                            rep.Append(table);
                        }
                        else {
                            errorCount++;
                            rep.AddError($"Не удалось найти дисциплину в учебном плане <b>{curriculum.Value.SourceFileName}</b>.");
                        }
                    }
                }
                else {
                    errorCount++;
                    rep.AddError("Не удалось найти учебные планы. Добавление УП осуществляется на вкладке <b>\"Учебные планы\"</b>.");
                    rep.AddError("Проверка РПД по УП невозможна.");
                }
                if (errorCount == 0) {
                    rep.AddDiv("Ошибок не обнаружено.");
                }
                rep.Append("</div></div>");

                toc.AddTocElement(rpd.DisciplineName ?? "?", anchor, errorCount);
            }
            rep.Append("</div>");
            toc.Append("</ul></div>");
            html.Append(toc).Append(rep).Append("</body></html>");

            htmlReport = html.ToString();
        }

        /// <summary>
        /// Найти учебные планы по формам обучения
        /// </summary>
        /// <param name="directionCode"></param>
        /// <param name="profile"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public static Dictionary<EFormOfStudy, Curriculum> FindCurricula(Rpd rpd) {
            if (!string.IsNullOrEmpty(rpd.DirectionCode) && !string.IsNullOrEmpty(rpd.Profile)) {  //} && !string.IsNullOrEmpty(rpd.Department)) {
                var items = m_curriculumDic.Values.Where(c => string.Compare(c.DirectionCode, rpd.DirectionCode, true) == 0 &&
                                                              string.Compare(c.Profile, rpd.Profile, true) == 0); // &&
                                                                                                                  //string.Compare(c.Department, rpd.Department, true) == 0);
                var dic = items.ToDictionary(x => x.FormOfStudy, x => x);
                return dic;
            }
            else {
                return null;
            }
        }

        /// <summary>
        /// Описание значения перечисления
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string GetDescription(this Enum value) {
            Type type = value.GetType();
            string name = Enum.GetName(type, value);
            if (name != null) {
                FieldInfo field = type.GetField(name);
                if (field != null) {
                    DescriptionAttribute attr = Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) as DescriptionAttribute;
                    if (attr != null) {
                        return attr.Description;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Генерация РПД
        /// </summary>
        /// <param name="curriculumGroup"></param>
        /// <param name="template"></param>
        /// <param name="targetDir"></param>
        /// <returns>список успешно созданных файлов</returns>
        public static List<string> GenerateRpdFiles(CurriculumGroup curriculumGroup, string rpdTemplate, string targetDir, string fileNameTemplate,
                                                    Action<int, CurriculumDiscipline> progressAction,
                                                    bool applyLoadedRpd,
                                                    out List<string> errors) {
            var files = new List<string>();
            errors = [];

            try {
                if (!File.Exists(rpdTemplate)) {
                    errors.Add($"Файл шаблона [{rpdTemplate}] не найден.");
                    return [];
                }
                if (!Directory.Exists(targetDir)) {
                    Directory.CreateDirectory(targetDir);
                }

                var i = 0;
                //цикл по дисциплинам
                foreach (var disc in curriculumGroup.CheckedDisciplines) {
                    try {
                        Rpd rpd = null;
                        if (applyLoadedRpd) {
                            rpd = FindRpd(disc);
                        }
                        i++;
                        progressAction?.Invoke(i, disc);
                        string fileName;
                        if (!string.IsNullOrEmpty(fileNameTemplate)) {
                            fileName = Regex.Replace(fileNameTemplate, @"({[^}]+})", (m) => disc.GetProperty(m.Value.Trim('{', '}'))?.ToString() ?? m.Value);
                        }
                        else {  //если шаблон имени файла не задан
                            fileName = $"РПД_{disc.Index}_{disc.Name}.docx";
                        }
                        //создание файла
                        var targetFile = Path.Combine(targetDir, fileName);
                        File.Copy(rpdTemplate, targetFile, true);

                        var eduWorks = new Dictionary<EFormOfStudy, EducationalWork>();
                        foreach (var curr in curriculumGroup.Curricula.Values) {
                            if (curr.Disciplines.TryGetValue(disc.Key, out CurriculumDiscipline currDisc)) {
                                eduWorks[curr.FormOfStudy] = currDisc.EducationalWork;
                            }
                        }

                        using (var docx = DocX.Load(targetFile)) {
                            //docx.InsertTableOfSummarys()
                            //этап 1. подстановка полей {...}
                            foreach (var par in docx.Paragraphs.ToList()) {
                                var replaceOptions = new FunctionReplaceTextOptions() {
                                    ContainerLocation = ReplaceTextContainer.All,
                                    FindPattern = "{[^}]+}",
                                    RegexMatchHandler = m => {
                                        var replaceValue = m;
                                        var propName = m.Trim('{', '}');
                                        object rpdProp = null;
                                        if (applyLoadedRpd && rpd != null) {
                                            rpdProp = rpd.GetProperty(propName);
                                        }
                                        if (rpdProp != null) {
                                            replaceValue = rpdProp.ToString();
                                        }
                                        else {
                                            var groupProp = curriculumGroup.GetProperty(propName);
                                            if (groupProp != null) {
                                                replaceValue = groupProp.ToString();
                                            }
                                            else {
                                                var discProp = disc.GetProperty(propName);
                                                if (discProp != null) {
                                                    replaceValue = discProp.ToString();
                                                    //if (string.IsNullOrEmpty(replaceValue)) {
                                                    //replaceValue = m;
                                                    //}
                                                }
                                                else {
                                                    if (TryProcessSpecialField(propName, docx, par, curriculumGroup, disc, rpd, out var specialValue)) {
                                                        replaceValue = specialValue;
                                                    }
                                                }
                                            }
                                        }
                                        return replaceValue;
                                    }
                                };
                                par.ReplaceText(replaceOptions);
                            }
                            var eduWorksIsOk = false;
                            var competenceTableIsOk = false;
                            var fullTimeTableIsOk = false;
                            var mixedTimeTableIsOk = false;
                            var partTimeTableIsOk = false;

                            //этап 2. заполнение/исправление таблиц
                            foreach (var table in docx.Tables.ToList()) {
                                if (!eduWorksIsOk) {
                                    //заполнение таблицы учебных работ
                                    eduWorksIsOk = TestForSummaryTableForEducationalWorks(table, eduWorks, PropertyAccess.Set /*установка значений в таблицу*/);
                                }
                                if (!competenceTableIsOk) {
                                    if (CompetenceMatrix.TestTable(table)) {
                                        RecreateTableOfCompetences(table, disc, false);
                                        competenceTableIsOk = true;
                                    }
                                }
                                if (applyLoadedRpd && rpd != null) {
                                    //проверка на таблицы учебных работ с темами
                                    var par = table.Paragraphs.FirstOrDefault();
                                    for (var j = 0; j < 5; j++) {
                                        par = par.PreviousParagraph;
                                        if (!fullTimeTableIsOk && m_eduWorkTableHeaders[EFormOfStudy.FullTime].IsMatch(par.Text)) {
                                            if (rpd.EducationalWorks.ContainsKey(EFormOfStudy.FullTime) && rpd.EducationalWorks[EFormOfStudy.FullTime].Table != null) {
                                                ResetTableIndentation(rpd.EducationalWorks[EFormOfStudy.FullTime].Table);
                                                var newTable = docx.AddTable(rpd.EducationalWorks[EFormOfStudy.FullTime].Table);
                                                par.InsertTableAfterSelf(newTable);
                                                table.Remove();
                                            }
                                            fullTimeTableIsOk = true;
                                            break;
                                        }
                                        if (!mixedTimeTableIsOk && m_eduWorkTableHeaders[EFormOfStudy.MixedTime].IsMatch(par.Text)) {
                                            if (rpd.EducationalWorks.ContainsKey(EFormOfStudy.MixedTime) && rpd.EducationalWorks[EFormOfStudy.MixedTime].Table != null) {
                                                ResetTableIndentation(rpd.EducationalWorks[EFormOfStudy.MixedTime].Table);
                                                var newTable = docx.AddTable(rpd.EducationalWorks[EFormOfStudy.MixedTime].Table);
                                                par.InsertTableAfterSelf(newTable);
                                                table.Remove();
                                            }
                                            mixedTimeTableIsOk = true;
                                            break;
                                        }
                                        if (!partTimeTableIsOk && m_eduWorkTableHeaders[EFormOfStudy.PartTime].IsMatch(par.Text)) {
                                            if (rpd.EducationalWorks.ContainsKey(EFormOfStudy.PartTime) && rpd.EducationalWorks[EFormOfStudy.PartTime].Table != null) {
                                                ResetTableIndentation(rpd.EducationalWorks[EFormOfStudy.PartTime].Table);
                                                var newTable = docx.AddTable(rpd.EducationalWorks[EFormOfStudy.PartTime].Table);
                                                par.InsertTableAfterSelf(newTable);
                                                table.Remove();
                                            }
                                            partTimeTableIsOk = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            docx.UpdateFields();
                            docx.Save();// ($"filled_{targetFile}");
                        }
                        files.Add(targetFile);
                    }
                    catch (Exception e) {
                        errors.Add($"Не удалось создать РПД для дисциплины [{disc.Name}]:\r\n{e.Message}\r\n{e.StackTrace}");
                    }
                }
            }
            catch (Exception ex) {
                errors.Add($"{ex.Message}\r\n{ex.StackTrace}");
            }

            return files;
        }

        /// <summary>
        /// Поиск РПД
        /// </summary>
        /// <param name="disc"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public static Rpd? FindRpd(CurriculumDiscipline disc) {
            var name = disc.Name.Replace('ё', 'е').ToLower();

            var rpd = m_rpdDic.Values.FirstOrDefault(d => d.DisciplineName.ToLower().Replace('ё', 'е').Equals(name));
            rpd ??= m_rpdDic.Values.FirstOrDefault(d => d.DisciplineName.ToLower().Replace('ё', 'е').StartsWith(name));
            rpd ??= m_rpdDic.Values.FirstOrDefault(d => name.StartsWith(d.DisciplineName.ToLower().Replace('ё', 'е')));

            return rpd; //m_rpdDic.Values.FirstOrDefault(r => r.DisciplineName.Equals(disc.Name, StringComparison.CurrentCultureIgnoreCase));
        }

        /// <summary>
        /// Попытка обработать специальное поле
        /// </summary>
        /// <param name="propName"></param>
        /// <param name="par"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private static bool TryProcessSpecialField(string propName, DocX docX, Paragraph par,
                                                   CurriculumGroup curriculumGroup, CurriculumDiscipline discipline,
                                                   Rpd rpd,
                                                   out string replaceValue) {
            var result = false;

            replaceValue = null;

            //проверка на поле типа {FullTime.TotalHours} для таблиц времен учебных работ
            var parts = propName.Split('.');
            if (parts.Length == 2) {
                if (Enum.TryParse(parts[0], true, out EFormOfStudy formOfStudy)) {
                    var curriculum = curriculumGroup.Curricula.Values.FirstOrDefault(c => c.FormOfStudy == formOfStudy);
                    if (curriculum != null) {
                        var disc = curriculum.FindDiscipline(discipline.Name);
                        if (disc != null) {
                            var propValue = disc.EducationalWork.GetProperty(parts[1]);
                            replaceValue = propValue?.ToString() ?? "-";
                            result = true;
                        }
                    }
                }
            }
            else {
                if (rpd != null) {
                    if (propName.Equals("Summary", StringComparison.CurrentCultureIgnoreCase) &&
                        (rpd.SummaryParagraphs?.Any() ?? false)) {
                        replaceValue = "";
                        //получение id стилей
                        var styleIdNormal = DocX.GetParagraphStyleIdFromStyleName(docX, "Normal"); //для сброса на стиль "Обычный" (Normal)
                        var currPar = par;
                        foreach (var sumPar in rpd.SummaryParagraphs) {
                            sumPar.SetLineSpacing(LineSpacingTypeAuto.None);
                            currPar = currPar.InsertParagraphAfterSelf(sumPar);
                            currPar.StyleId = styleIdNormal;
                            currPar.FontSize(14);
                            currPar.ShadingPattern(new ShadingPattern() { Fill = Color.Yellow }, ShadingType.Paragraph);
                        }
                        result = true;
                    }
                    if (propName.Equals("Questions", StringComparison.CurrentCultureIgnoreCase) &&
                        (rpd.QuestionList?.Any() ?? false)) {
                        replaceValue = "";
                        //var numberedList = docX.AddList(new ListOptions());
                        //numberedList.
                        var currPar = par;
                        var num = 0;
                        var shadingPattern = new ShadingPattern() { Fill = Color.Yellow };
                        foreach (var questionText in rpd.QuestionList) {
                            num++;
                            currPar = currPar.InsertParagraphAfterSelf("");
                            currPar.IndentationFirstLine = 35;
                            currPar.Append($"{num}. {questionText}").FontSize(14).ShadingPattern(shadingPattern, ShadingType.Paragraph);
                        }
                        result = true;
                    }
                    if (propName.Equals("BaseReferences", StringComparison.CurrentCultureIgnoreCase) &&
                        (rpd.ReferencesBase?.Any() ?? false)) {
                        replaceValue = "";
                        var currPar = par;
                        var num = 0;
                        var shadingPattern = new ShadingPattern() { Fill = Color.Yellow };
                        foreach (var item in rpd.ReferencesBase) {
                            num++;
                            currPar = currPar.InsertParagraphAfterSelf("");
                            currPar.IndentationFirstLine = 35;
                            currPar.Append($"{num}. {item}").FontSize(14).ShadingPattern(shadingPattern, ShadingType.Paragraph);
                        }
                        result = true;
                    }
                    if (propName.Equals("ExtraReferences", StringComparison.CurrentCultureIgnoreCase) &&
                        (rpd.ReferencesBase?.Any() ?? false)) {
                        replaceValue = "";
                        var currPar = par;
                        var num = 0;
                        var shadingPattern = new ShadingPattern() { Fill = Color.Yellow };
                        foreach (var item in rpd.ReferencesExtra) {
                            num++;
                            currPar = currPar.InsertParagraphAfterSelf("");
                            currPar.IndentationFirstLine = 35;
                            currPar.Append($"{num}. {item}").FontSize(14).ShadingPattern(shadingPattern, ShadingType.Paragraph);
                        }
                        result = true;
                    }
                }

                if (propName.Equals("TOC", StringComparison.CurrentCultureIgnoreCase)) {
                    //                    var tocSwitches = new Dictionary<TableOfSummarysSwitches, string>();
                    ////{
                    ////          { TableOfSummarysSwitches.O, "1-3"},
                    ////          { TableOfSummarysSwitches.U, ""},
                    ////          { TableOfSummarysSwitches.Z, ""},
                    ////          { TableOfSummarysSwitches.H, ""},
                    ////        };
                    //                    TableOfSummarys toc = docX.InsertTableOfSummarys(par, "", tocSwitches);
                    //                    par.FontSize(14);
                    //                    replaceValue = "";
                }
            }
            //switch (propName) {
            //    ////case "TableCompetences":
            //    ////    var newTable = par.InsertTableAfterSelf(1, 3);
            //    ////    RecreateTableOfCompetences(newTable, discipline);
            //    ////    result = true;
            //    ////    break;
            //    //case "TableEducationWorks":
            //    //    par.IndentationFirstLine = 0.1f;
            //    //    var newTable2 = par.InsertTableAfterSelf(8, 4);
            //    //    CreateTableEducationWorks(newTable2, discipline);
            //    //    result = true;
            //    //    break;
            //    //case "DisciplineTypeDescription":
            //    //    var text = "";
            //    //    if (discipline.Type == EDisciplineType.Required) {
            //    //        text = "обязательным дисциплинам";
            //    //    }
            //    //    else if (discipline.Type == EDisciplineType.ByChoice) {
            //    //        text = "дисциплинам по выбору части, формируемой участниками образовательных отношений";
            //    //    }
            //    //    break;
            //    default:
            //        break;
            //}

            return result;
        }


        /// <summary>
        /// Получить или установить значение свойства из/в параграф(а)
        /// </summary>
        /// <param name="par"></param>
        /// <param name="propMatchCheck"></param>
        /// <param name="testText"></param>
        /// <param name="getValue"></param>
        /// <param name="value"></param>
        /// <param name="prop"></param>
        static void GetOrSetPropValue(Paragraph par, Regex propMatchCheck, string testText, PropertyAccess propAccess, TypeAccessor typeAccessor,
                                      object obj, string propName, Type propType) {
            if (propMatchCheck.IsMatch(testText)) {
                if (propAccess == PropertyAccess.Get) {
                    var text = par.Text;
                    if (propType == typeof(int)) {
                        if (int.TryParse(text, out int intVal)) {
                            typeAccessor[obj, propName] = intVal;
                        }
                    }
                }
                else {
                    var replaceText = typeAccessor[obj, propName]?.ToString() ?? "-";
                    if (!string.IsNullOrEmpty(replaceText)) {
                        var replaceTextOptions = new FunctionReplaceTextOptions() {
                            FindPattern = @"^.*$",
                            ContainerLocation = ReplaceTextContainer.All,
                            RegexMatchHandler = m => replaceText
                        };
                        par.ReplaceText(replaceTextOptions);
                    }
                }
            }
        }

        /// <summary>
        /// Проверка с чтением или заполнением полей таблицы на учебные работы по формам обучения
        /// </summary>
        /// <param name="table"></param>
        /// <param name="eduWorks">словарь либо для получения, либо для простановки времен учебной работы</param>
        /// <param name="propAccess">режим работы: получить или установить значения</param>
        public static bool TestForSummaryTableForEducationalWorks(Table table, Dictionary<EFormOfStudy, EducationalWork> eduWorks, PropertyAccess propAccess) {
            var result = false;

            if (table.RowCount > 5 && table.ColumnCount >= 2) {
                eduWorks ??= [];

                var headerRow = table.Rows[0];
                var cellText = headerRow.Cells[0].GetText();
                //проверка левой верхней ячейки
                if (m_regexEduWorkType.IsMatch(cellText) ||
                    string.IsNullOrWhiteSpace(cellText) /* встречаются РПД, где в этой ячейке таблицы пусто */) {
                    Dictionary<EFormOfStudy, int> formColIdx = new() {
                        { EFormOfStudy.FullTime, -1 },
                        { EFormOfStudy.MixedTime, -1 },
                        { EFormOfStudy.PartTime, -1 }
                    };
                    for (var colIdx = 1; colIdx < headerRow.Cells.Count; colIdx++) {
                        var text = headerRow.Cells[colIdx].GetText();
                        if (!string.IsNullOrEmpty(text)) {
                            if (m_regexFormFullTime.IsMatch(text)) {
                                formColIdx[EFormOfStudy.FullTime] = colIdx;
                                if (propAccess == PropertyAccess.Get) eduWorks[EFormOfStudy.FullTime] = new();
                            }
                            else if (m_regexFormMixedTime.IsMatch(text)) {
                                formColIdx[EFormOfStudy.MixedTime] = colIdx;
                                if (propAccess == PropertyAccess.Get) eduWorks[EFormOfStudy.MixedTime] = new();
                            }
                            else if (m_regexFormPartTime.IsMatch(text)) {
                                formColIdx[EFormOfStudy.PartTime] = colIdx;
                                if (propAccess == PropertyAccess.Get) eduWorks[EFormOfStudy.PartTime] = new();
                            }
                        }
                    }
                    //если колонки форм обучения обнаружены
                    if (formColIdx.Any(x => x.Value >= 0)) {
                        foreach (var item in formColIdx) {
                            if (eduWorks.TryGetValue(item.Key, out var eduWork)) {
                                //выявление рядов
                                for (var rowIdx = 1; rowIdx < table.RowCount; rowIdx++) {
                                    var par = table.Rows[rowIdx].Cells[item.Value].Paragraphs?.FirstOrDefault();
                                    var text = par.Text;

                                    var header = table.Rows[rowIdx].Cells[0].GetText();
                                    GetOrSetPropValue(par, m_regexTotalHours, header, propAccess, EducationalWork.TypeAccessor, eduWork, nameof(eduWork.TotalHours), typeof(int));
                                    GetOrSetPropValue(par, m_regexContactHours, header, propAccess, EducationalWork.TypeAccessor, eduWork, nameof(eduWork.ContactWorkHours), typeof(int));
                                    GetOrSetPropValue(par, m_regexControlHours, header, propAccess, EducationalWork.TypeAccessor, eduWork, nameof(eduWork.ControlHours), typeof(int));
                                    GetOrSetPropValue(par, m_regexLectureHours, header, propAccess, EducationalWork.TypeAccessor, eduWork, nameof(eduWork.LectureHours), typeof(int));
                                    GetOrSetPropValue(par, m_regexLabHours, header, propAccess, EducationalWork.TypeAccessor, eduWork, nameof(eduWork.LabHours), typeof(int));
                                    GetOrSetPropValue(par, m_regexPracticalHours, header, propAccess, EducationalWork.TypeAccessor, eduWork, nameof(eduWork.PracticalHours), typeof(int));
                                    GetOrSetPropValue(par, m_regexSelfStudyHours, header, propAccess, EducationalWork.TypeAccessor, eduWork, nameof(eduWork.SelfStudyHours), typeof(int));
                                    //форма итогового контроля
                                    if (m_regexControlForm.IsMatch(header)) {
                                        if (propAccess == PropertyAccess.Get) {
                                            if (m_regexControlExam.IsMatch(text)) {
                                                eduWork.ControlForm = EControlForm.Exam;
                                            }
                                            else if (m_regexControlTestGrade.IsMatch(text)) {
                                                eduWork.ControlForm = EControlForm.TestWithAGrade;
                                            }
                                            else if (m_regexControlTest.IsMatch(text)) {
                                                eduWork.ControlForm = EControlForm.Test;
                                            }
                                            else {
                                                eduWork.ControlForm = EControlForm.Unknown;
                                            }
                                        }
                                        else {
                                            var replaceOptions = new FunctionReplaceTextOptions() {
                                                ContainerLocation = ReplaceTextContainer.All,
                                                FindPattern = @"^.*$",
                                                RegexMatchHandler = m => eduWork.ControlForm.GetDescription().ToLower()
                                            };
                                            par.ReplaceText(replaceOptions);
                                        }
                                    }
                                }
                            }
                        }
                        result = true;
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Создание/заполнение таблицы учебных работ
        /// </summary>
        /// <param name="table"></param>
        /// <param name="discipline"></param>
        //private static void CreateTableEducationWorks(Table table, CurriculumDiscipline discipline) {
        //    var widths = new double[4] { 8.46 * 28.35, 2.5 * 28.35, 2.66 * 28.35, 2.79 * 28.35 };
        //    //заголовок
        //    var header0 = table.Rows[0].Cells[0].Paragraphs.FirstOrDefault();
        //    table.Rows[0].Cells[0].Width = widths[0];
        //    header0.InsertText("Виды учебной работы", false, formatting: new Formatting() { Bold = true });
        //    header0.Alignment = Alignment.center;
        //    header0.IndentationFirstLine = 0.1f;

        //    var header1 = table.Rows[0].Cells[1].Paragraphs.FirstOrDefault();
        //    table.Rows[0].Cells[1].Width = widths[1];
        //    header1.InsertText("очная форма\r\nобучения", formatting: new Formatting() { Bold = false });
        //    header1.Alignment = Alignment.center;
        //    header1.IndentationFirstLine = 0.1f;

        //    var header2 = table.Rows[0].Cells[2].Paragraphs.FirstOrDefault();
        //    table.Rows[0].Cells[2].Width = widths[2];
        //    header2.InsertText("очно-заочная\r\nобучения", formatting: new Formatting() { Bold = false });
        //    header2.Alignment = Alignment.center;
        //    header2.IndentationFirstLine = 0.1f;

        //    var header3 = table.Rows[0].Cells[3].Paragraphs.FirstOrDefault();
        //    table.Rows[0].Cells[3].Width = widths[3];
        //    header3.InsertText("заочная форма\r\nобучения", formatting: new Formatting() { Bold = false });
        //    header3.Alignment = Alignment.center;
        //    header3.IndentationFirstLine = 0.1f;
        //}

        /// <summary>
        /// Пересоздание таблицы с компетенциями (опциональное пересоздание заголовка, остальные ряды формируются заново)
        /// </summary>
        /// <param name="table">пустая таблицы (1 ряд, 3 колонки)</param>
        /// <param name="discipline"></param>
        static bool RecreateTableOfCompetences(Table table, CurriculumDiscipline discipline, bool recreateHeaders) {
            var result = false;

            //очистка таблицы
            if (table.ColumnCount < 3) {
                return false;
            }
            var topRowIdxToRemove = recreateHeaders ? 0 : 1;
            for (var rowIdx = table.RowCount - 1; rowIdx >= topRowIdxToRemove; rowIdx--) {
                table.RemoveRow(rowIdx);
            }

            if (table.ColumnCount > 3) {
                var t = 0;
                //for (var colIdx = table.ColumnCount - 1; colIdx >= 3; colIdx--) {
                //    table.RemoveColumn(colIdx);
                //}
            }

            if (recreateHeaders) {
                //table.InsertRow();
                //table.InsertColumn(0, false);
                //table.InsertColumn();
                //table.InsertColumn();
                //заголовок
                var header0 = table.Rows[0].Cells[0].Paragraphs.FirstOrDefault();
                header0.InsertText("Код и наименование", false, formatting: new Formatting() { Bold = true });
                header0.Alignment = Alignment.center;
                header0.IndentationFirstLine = 0.1f;
                var header02 = header0.InsertParagraphAfterSelf("компетенций", false, formatting: new Formatting() { Bold = true });
                header02.Alignment = Alignment.center;
                header02.IndentationFirstLine = 0.1f;

                var header1 = table.Rows[0].Cells[1].Paragraphs.FirstOrDefault();
                header1.InsertText("Коды и индикаторы", formatting: new Formatting() { Bold = true });
                header1.Alignment = Alignment.center;
                header1.IndentationFirstLine = 0.1f;
                var header12 = header1.InsertParagraphAfterSelf("достижения компетенций", false, formatting: new Formatting() { Bold = true });
                header12.Alignment = Alignment.center;
                header12.IndentationFirstLine = 0.1f;

                var header2 = table.Rows[0].Cells[2].Paragraphs.FirstOrDefault();
                header2.InsertText("Коды и результаты обучения", formatting: new Formatting() { Bold = true });
                header2.Alignment = Alignment.center;
                header2.IndentationFirstLine = 0.1f;
            }

            //ряды матрицы
            var appliedIndicatorCount = 0;
            var items = CompetenceMatrix.GetItems(discipline.CompetenceList); //отбор строк матрицы по списку индикаторов
            foreach (var item in items) {
                var competenceStartRow = table.RowCount;

                foreach (var achi in item.Achievements) {
                    if (!discipline.CompetenceList.Contains(achi.Code)) { //доп. фильтрация только по нужным индикаторам
                        continue;
                    }
                    appliedIndicatorCount++;

                    var achiStartRow = table.RowCount;
                    foreach (var res in achi.Results) {
                        var resRowIdx = table.RowCount;
                        var newRow = table.InsertRow();
                        if (resRowIdx == competenceStartRow) {
                            //ячейка компетенции
                            var cellCompetence = newRow.Cells[0].Paragraphs.FirstOrDefault();
                            //table.MergeCellsInColumn(0, firstRowIdx, firstRowIdx + rowSpan - 1);
                            cellCompetence.InsertText($"{item.Code}.", formatting: new Formatting() { Bold = true });
                            cellCompetence.InsertText($" {item.Title}", formatting: new Formatting() { Bold = false });
                            cellCompetence.IndentationFirstLine = 0.1f;
                        }
                        if (resRowIdx == achiStartRow) {
                            //ячейка индикатора
                            var cellIndicator = newRow.Cells[1].Paragraphs.FirstOrDefault();
                            cellIndicator.InsertText($"{achi.Code}. {achi.Indicator}", formatting: new Formatting() { Bold = false });
                            cellIndicator.IndentationFirstLine = 0.1f;
                        }

                        //ячейка результата
                        var cellResult = newRow.Cells[2].Paragraphs.FirstOrDefault();
                        cellResult.InsertText($"{res.Code}:", formatting: new Formatting() { Bold = false });
                        cellResult.IndentationFirstLine = 0.1f;
                        var cellResult2 = cellResult.InsertParagraphAfterSelf($"{res.Description}", false, formatting: new Formatting() { Bold = false });
                        cellResult2.IndentationFirstLine = 0.1f;
                    }
                    if (table.RowCount - 1 > achiStartRow) {
                        table.MergeCellsInColumn(1, achiStartRow, table.RowCount - 1);
                    }
                }
                if (table.RowCount - 1 > competenceStartRow) {
                    table.MergeCellsInColumn(0, competenceStartRow, table.RowCount - 1);
                }
            }
            result = appliedIndicatorCount == discipline.CompetenceList.Count;

            return result;
        }

        /// <summary>
        /// Сохранить конфиг приложения
        /// </summary>
        static public void SaveConfig() {
            var configFile = Path.Combine(Environment.CurrentDirectory, $"{Assembly.GetExecutingAssembly().GetName().Name}.config.json");
            if (m_config != null) {
                var json = JsonSerializer.Serialize(m_config, m_jsonOptions);
                File.WriteAllText(configFile, json, Encoding.UTF8);
            }
        }

        /// <summary>
        /// Сериализация объекта
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string SerializeObj(object obj) {
            return JsonSerializer.Serialize(obj, m_jsonOptions);
        }

        /// <summary>
        /// Десериализация JSON в объект
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="json"></param>
        /// <returns></returns>
        public static T Deserialize<T>(string json) {
            return JsonSerializer.Deserialize<T>(json, m_jsonOptions);
        }

        /// <summary>
        /// Загрузить конфиг
        /// </summary>
        static public void LoadConfig() {
            var configFile = Path.Combine(Environment.CurrentDirectory, $"{Assembly.GetExecutingAssembly().GetName().Name}.config.json");
            if (File.Exists(configFile)) {
                var json = File.ReadAllText(configFile);
                m_config = JsonSerializer.Deserialize<Config>(json, m_jsonOptions);
            }
            else {
                m_config = new();
            }

            if (!(m_config.CurriculumDisciplineParseItems?.Any() ?? false)) {
                m_config.CurriculumDisciplineParseItems = CurriculumDisciplineReader.DefaultHeaders;
            }
            if (!(m_config.Departments?.Any() ?? false)) {
                m_config.Departments = Department.DefaultDepartments;
            }
            if (!(m_config.RpdFixDocPropertyList?.Any() ?? false)) {
                m_config.RpdFixDocPropertyList = DocProperty.DefaultProperties;
            }
        }

        /// <summary>
        /// Режим фикса РПД
        /// </summary>
        /// <param name="rpdList"></param>
        internal static void FixRpdFiles(List<Rpd> rpdList, string targetDir, out string htmlReport) {
            var html = new StringBuilder("<html><body><h2>Отчёт по исправлению РПД</h2>");
            DocX templateDocx = null;
            DocxChapters templateChapters = null;

            try {
                var fixByTemplate = false;
                if (Config.RpdFixByTemplate) {
                    if (File.Exists(Config.RpdFixTemplateFileName)) {
                        templateDocx = DocX.Load(Config.RpdFixTemplateFileName);
                        templateChapters = ScanDocxForChapters(templateDocx);
                        fixByTemplate = true;
                    }
                }

                html.Append("<div><b>Режим работы:</b></div><ul>");
                if (Config.RpdFixTableOfCompetences) {
                    html.Append("<li>Исправление таблицы компетенций</li>");
                }
                if (Config.RpdFixTableOfEduWorks) {
                    html.Append("<li>Исправление таблицы учебных работ</li>");
                }
                if (Config.RpdFixFillEduWorkTables) {
                    html.Append("<li>Заполнение таблиц учебных работ для форм обучения</li>");
                }
                var findAndReplaceItems = Config.RpdFindAndReplaceItems?.Where(i => i.IsChecked).ToList();
                if (findAndReplaceItems != null && findAndReplaceItems.Any()) {
                    var tdStyle = "style='border: 1px solid;'";
                    html.Append($"<li><table {tdStyle}><tr><th {tdStyle}><b>Найти</b></th><th {tdStyle}><b>Заменить на</b></th></tr>");
                    foreach (var item in findAndReplaceItems) {
                        html.Append($"<tr><td {tdStyle}>{item.FindPattern}</td><td {tdStyle}>{item.ReplacePattern}</td></tr>");
                    }
                    html.Append("</table></li>");
                }
                html.Append("</ul>");

                foreach (var rpd in rpdList) {
                    html.Append("<p />");
                    if (File.Exists(rpd.SourceFileName)) {
                        html.Append($"<div>Исходный файл: <b>{rpd.SourceFileName}</b></div>");
                        html.Append($"<div>Дисциплина: <b>{rpd.DisciplineName}</b></div>");

                        var fixCompetences = Config.RpdFixTableOfCompetences;
                        var discipline = FindDiscipline(rpd);
                        if (discipline == null) {
                            fixCompetences = false;
                            html.Append($"<div style='color: red'>Не удалось найти дисциплину [{rpd.DisciplineName}] в загруженных учебных планах</div>");
                        }
                        var fixEduWorksSummary = Config.RpdFixTableOfEduWorks;
                        var eduWorks = GetEducationWorks(rpd, out _);
                        //if (!(eduWorks?.Any() ?? false)) {
                        //    fixEduWorks = false;
                        //    html.Append($"<div style='color: red'>Не удалось найти учебные работы для дисциплины [{rpd.DisciplineName}] в загруженных учебных планах</div>");
                        //}
                        if (rpd.FormsOfStudy.Count != eduWorks.Count) {
                            fixEduWorksSummary = false;
                            foreach (var form in rpd.FormsOfStudy) {
                                if (!eduWorks.ContainsKey(form)) {
                                    html.Append($"<div style='color: red'>Учебный план для формы обучения [{form.GetDescription()}] не загружен</div>");
                                }
                            }
                        }
                        var fixEduWorkTables = Config.RpdFixFillEduWorkTables;
                        var eduWorkTableIsFixed = rpd.FormsOfStudy.ToDictionary(x => x, x => false);

                        var eduSummaryTableIsFixed = false; //флаг, что в процессе работы была исправлена сводная таблица учебных работ

                        var setDocProperties = Config.RpdFixDocPropertyList?.Where(i => i.IsChecked).ToList();

                        using (var docx = DocX.Load(rpd.SourceFileName)) {
                            if (setDocProperties?.Any() ?? false) {
                                foreach (var item in setDocProperties) {
                                    ///docx.CoreProperties[item.Name] = item.Value;
                                    docx.AddCoreProperty(item.Name, item.Value);
                                }
                            }

                            foreach (var table in docx.Tables) {
                                var backup = table.Xml;
                                if (fixCompetences && CompetenceMatrix.TestTable(table)) {
                                    if (RecreateTableOfCompetences(table, discipline, false)) {
                                        html.Append("<div style='color: green'>Таблица компетенций сформирована по матрице компетенций.</div>");
                                    }
                                    else {
                                        html.Append("<div style='color: red'>Не удалось сформировать обновленную таблицу компетенций.</div>");
                                        table.Xml = backup;
                                    }
                                }
                                if (fixEduWorksSummary && !eduSummaryTableIsFixed) {
                                    eduSummaryTableIsFixed |= TestForSummaryTableForEducationalWorks(table, eduWorks, PropertyAccess.Set /*установка значений в таблицу*/);
                                }
                                if (fixEduWorkTables) {
                                    //фикс-заполнение таблицы учебных работ для формы обучения
                                    if (TestForEduWorkTable(table, rpd, PropertyAccess.Set, out var formOfStudy)) {
                                        eduWorkTableIsFixed[formOfStudy] = true;
                                    }
                                }
                            }
                            if (fixEduWorksSummary) {
                                if (eduSummaryTableIsFixed) {
                                    html.Append("<div style='color: green'>Сводная таблица учебных работ заполнена по учебным планам.</div>");
                                }
                                else {
                                    html.Append("<div style='color: red'>Не удалось сформировать сводную таблицу учебных работ.</div>");
                                    //table.Xml = backup;
                                }
                            }
                            if (fixEduWorkTables) {
                                foreach (var item in eduWorkTableIsFixed) {
                                    if (item.Value) {
                                        html.Append($"<div style='color: green'>Таблица учебных работ для формы обучения [{item.Key.GetDescription()}] успешно заполнена.</div>");
                                    }
                                    else {
                                        html.Append($"<div style='color: red'>Таблицу учебных работ для формы обучения [{item.Key.GetDescription()}] не удалось заполнить.</div>");
                                    }
                                }
                            }

                            //обработка "найти и заменить"
                            var replaceCount = 0;
                            if (findAndReplaceItems?.Any() ?? false) {
                                foreach (var par in docx.Paragraphs) {
                                    foreach (var findItem in findAndReplaceItems) {
                                        var replaceOptions = new FunctionReplaceTextOptions() {
                                            FindPattern = findItem.FindPattern,
                                            ContainerLocation = ReplaceTextContainer.All,
                                            StopAfterOneReplacement = false,
                                            RegexMatchHandler = m => findItem.ReplacePattern,
                                            RegExOptions = RegexOptions.IgnoreCase
                                        };
                                        if (par.ReplaceText(replaceOptions)) {
                                            replaceCount++;
                                        }
                                    }
                                }
                                html.Append($"<div>Осуществлено замен в тексте: {replaceCount}</div>");
                            }

                            //фикс разделов по шаблону
                            if (fixByTemplate) {
                                var docxSections = ScanDocxForChapters(docx);
                            }

                            var fileName = Path.GetFileName(rpd.SourceFileName);
                            var newFileName = Path.Combine(targetDir, fileName);
                            if (!Directory.Exists(targetDir)) {
                                Directory.CreateDirectory(targetDir);
                            }
                            docx.SaveAs(newFileName);
                            html.Append($"<div>Итоговый файл: <b>{newFileName}</b></div>");
                        }
                    }
                    else {
                        html.Append($"<div style='color:red'>Файл <b>{rpd.SourceFileName}</b> не найден</div>");
                    }
                }
            }
            catch (Exception ex) {
                html.Append($"<div>{ex.Message}</div>");
                html.Append($"<div>{ex.StackTrace}</div>");
            }
            finally {
                templateDocx?.Dispose();
                html.Append("</body></html>");
            }

            htmlReport = html.ToString();
        }

        /// <summary>
        /// Получить коллекцию учебных работ для указанного РПД
        /// </summary>
        /// <param name="rpd"></param>
        /// <returns></returns>
        public static Dictionary<EFormOfStudy, EducationalWork> GetEducationWorks(Rpd rpd, out Dictionary<EFormOfStudy, Curriculum> curricula) {
            curricula = FindCurricula(rpd);
            var eduWorks = new Dictionary<EFormOfStudy, EducationalWork>();
            if (curricula != null) {
                foreach (var curr in curricula.Values) {
                    var disc = curr.FindDiscipline(rpd.DisciplineName);
                    if (disc != null) {
                        eduWorks[curr.FormOfStudy] = disc.EducationalWork;
                    }
                }
            }
            return eduWorks;
        }

        /// <summary>
        /// Найти дисциплину для указанного РПД
        /// </summary>
        /// <param name="rpd"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public static CurriculumDiscipline FindDiscipline(Rpd rpd) {
            CurriculumDiscipline discipline = null;
            var curricula = FindCurricula(rpd);
            if (curricula != null) {
                discipline = curricula.Values.FirstOrDefault()?.FindDiscipline(rpd.DisciplineName);
                //var curriculum = curricula.Values.Select(c => c.FormOfStudy == rpd.FormsOfStudy)
            }

            return discipline;
        }


        /// <summary>
        /// Сканирование документа на разделы
        /// </summary>
        /// <param name="docX"></param>
        internal static DocxChapters ScanDocxForChapters(DocX docX) {
            var docxSections = new DocxChapters() {
                Chapters = []
            };

            DocxChapter section = null;
            foreach (var par in docX.Paragraphs) {
                bool skipPar = false;
                var magicText = par.MagicText?.FirstOrDefault();
                if (magicText != null && magicText.formatting?.Bold == true) {
                    var m = m_regexRpdChapter.Match(magicText.text);
                    if (m.Success) {
                        section = new() {
                            Number = m.Groups[1].Value,
                            Text = m.Groups[2].Value,
                            Paragraphs = []
                        };
                        docxSections.Chapters.Add(section);
                        skipPar = true;
                    }
                }
                if (!skipPar && section != null) {
                    section.Paragraphs.Add(par);
                }
            }

            return docxSections;
        }

        /// <summary>
        /// Сброс отступов в таблице
        /// </summary>
        /// <param name="table"></param>
        internal static void ResetTableIndentation(Table table) {
            table.Paragraphs.ForEach(p => {
                p.IndentationFirstLine = 0.1f; // .1f;
                p.IndentationBefore = 0.1f;
                p.IndentationHanging = 0.1f;
                p.IndentationAfter = 0.1f;
            });
        }

        /// <summary>
        /// Проверка на совпадение имён дисциплин
        /// </summary>
        /// <param name="name1"></param>
        /// <param name="name2"></param>
        /// <returns></returns>
        internal static bool IsDisciplineNamesEquals(string name1, string name2) {
            name1 = name1.Trim().Replace('ё', 'е').ToLower();
            name2 = name2.Trim().Replace('ё', 'е').ToLower();

            return name1.Equals(name2) || name1.StartsWith(name2) || name2.StartsWith(name1);
        }

        //static public void ForEach<T>(this T[] arr, Action<T> action) {
        //    for (int i = 0; i < arr.Length; i++) {
        //        action.Invoke(arr[i]);
        //    }
        //}

        /// <summary>
        /// Функция для распределения времени
        /// </summary>
        /// <param name="totalValue"></param>
        /// <param name="count"></param>
        /// <param name="luckyNumbers"></param>
        /// <returns></returns>
        static int[] SplitTime(int total, int count, HashSet<int> luckyNumbers) {
            var items = new int[count];

            var extraTime = 0;

            double oneItemTime = (double)total / count;
            if (oneItemTime >= 2) { //хватает времени на пару для всех топиков?
                if (oneItemTime % 2 == 0) { //удалось поделить на всех поровну?
                    for (var i = 0; i < count; i++) {
                        items[i] = Convert.ToInt32(oneItemTime);
                    }
                }
                else { //поровну не хватило, надо делить
                    //приводим значение к ближайшему четному в меньшую сторону
                    var fixedTime = (int)oneItemTime - 1;
                    for (int i = 0; i < items.Length; i++) { items[i] = fixedTime; }
                    extraTime = total - fixedTime * count;   //время для распределения
                }
            }
            else {
                extraTime = total;
            }
            if (extraTime > 0) {
                var luckyTopicCount = extraTime >> 1;               //кол-во счастливых топиков, которые получат доп. время
                //var unluckyTopicCount = count - luckyTopicCount;    //оставшиеся - несчастные
                //выбираем случайным образом счастливые топики с учетом уже выбранных номеров, переданных функции
                //var numbers = luckyNumbers?.Take(luckyTopicCount).ToHashSet() ?? new();
                var rand = new Random();
                while (luckyNumbers.Count < luckyTopicCount) {
                    while (true) {
                        var num = rand.Next(count);
                        if (luckyNumbers.Add(num)) {
                            break;
                        }
                    }
                }
                foreach (var num in luckyNumbers.Take(luckyTopicCount)) {
                    items[num] += 2;
                }
            }

            return items;
        }

        /// <summary>
        /// Поделить время обучения
        /// </summary>
        /// <param name="topicCount"></param>
        /// <param name="eduWork"></param>
        /// <returns></returns>
        static (int total, int contact, int lecture, int practical, int selfStudy)[] SplitEduTime(int topicCount, EducationalWork eduWork) {
            var topics = new (int total, int contact, int lecture, int practical, int selfStudy)[topicCount];

            //делим лекционное время
            var lectureHours = eduWork.LectureHours;
            double oneLectureTime = (double)lectureHours / topicCount;
            if (oneLectureTime >= 2) { //хватает времени на пару для всех топиков?
                if (oneLectureTime % 2 == 0) { //удалось поделить на всех поровну?
                    for (var topic = 0; topic < topicCount; topic++) {
                        topics[topic].lecture = Convert.ToInt32(oneLectureTime);
                    }
                }
                else { //поровну не хватило, надо делить
                    //приводим значение к ближайшему четному в меньшую сторону
                    var lecTime = (int)oneLectureTime - 1;
                    for (int i = 0; i < topics.Length; i++) { topics[i].lecture = lecTime; }
                    var lecExtraTime = lectureHours - lecTime * topicCount; //время для распределения
                    var luckyTopicCount = lecExtraTime >> 1;                //кол-во счастливых топиков, которые получат доп. время
                    var unluckyTopicCount  = topicCount - luckyTopicCount;  //оставшиеся - несчастные
                    //выбираем случайным образом счастливые топики
                    var luckyNumbers = new HashSet<int>();
                    var rand = new Random();
                    while (luckyNumbers.Count < luckyTopicCount) {
                        while (true) {
                            var num = rand.Next(topicCount);
                            if (luckyNumbers.Add(num)) {
                                break;
                            }
                        }
                    }
                    //topics.ForEach(t => t.lecture = 6);
                    foreach (var num in luckyNumbers) {
                        topics[num].lecture += 2;
                    }
                    var t = 0;
                    //while (true) {
                    //    for (var topic = 0; topic < topicCount - 1; topic += 2) {
                    //        topics[topic + 1].lecture -= 1;
                    //    }
                    //    topics[t].lecture--;
                    //    //topics
                    //}
                }
            }
            else { //не хватает времени на все топики

            }

            return topics;
        }
        
        /// <summary>
        /// Заполнение таблицы учебных работ по заданной форме обучения
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rpd"></param>
        /// <param name="formOfStudy"></param>
        static void FillEduWorkTable(Table table, Rpd rpd, EFormOfStudy formOfStudy) {
            var curricula = FindCurricula(rpd);
            if (curricula?.Any() ?? false) {
                if (table.ColumnCount == 8) {
                    var topicCount = table.RowCount - 5;
                    //распределяем время (время должно быть четным)
                    var topics = SplitEduTime(topicCount, rpd.EducationalWorks[formOfStudy]); // new (int total, int contact, int lecture, int practical, int selfStudy)[topicCount];
                    
                    for (int row = 3; row < table.RowCount - 2; row++) {
                        var cellTopic = table.Rows[row].Cells[0];
                        var cellTimeTotal = table.Rows[row].Cells[1];
                        var cellTimeContactTotal = table.Rows[row].Cells[2];
                        var cellTimeLectures = table.Rows[row].Cells[3];
                        var cellTimePractical = table.Rows[row].Cells[4];
                        var cellTimeSelfStudy = table.Rows[row].Cells[5];
                        var cellEvalTools = table.Rows[row].Cells[6];
                        var cellResults = table.Rows[row].Cells[7];
                    }
                }
            }
        }

        /// <summary>
        /// Проверка на таблицу учебных работ по форме обучения
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rpd"></param>
        /// <returns></returns>
        internal static bool TestForEduWorkTable(Table table, Rpd rpd, PropertyAccess propAccess, out EFormOfStudy formOfStudy) {
            var result = false;
            formOfStudy = EFormOfStudy.Unknown;

            //проверка на таблицы учебных работ с темами
            var par = table.Paragraphs.FirstOrDefault();
            for (var i = 0; i < 5; i++) {
                par = par?.PreviousParagraph;
                if (par == null) {
                    break;
                }
                if (string.IsNullOrWhiteSpace(par.Text)) {
                    continue;
                }
                foreach (EFormOfStudy form in Enum.GetValues(typeof(EFormOfStudy))) {
                    if (m_eduWorkTableHeaders.TryGetValue(form, out var regex) &&
                        rpd.EducationalWorks.ContainsKey(form) &&
                        regex.IsMatch(par.Text)) {
                        if (propAccess == PropertyAccess.Get) {
                            if (rpd.EducationalWorks[form].Table == null) {
                                rpd.EducationalWorks[form].Table = table;
                            }
                        }
                        else {
                            //простановка времени по темам
                            FillEduWorkTable(table, rpd, form);
                        }
                        result = true;
                        break;
                    }
                }
                if (result) break;
                //if (rpd.EducationalWorks.ContainsKey(EFormOfStudy.FullTime) && 
                //    rpd.EducationalWorks[EFormOfStudy.FullTime].Table == null && 
                //    m_regexFullTimeTable.IsMatch(par.Text)) {
                //    rpd.EducationalWorks[EFormOfStudy.FullTime].Table = table;
                //    result = true;
                //    break;
                //}
                //if (rpd.EducationalWorks.ContainsKey(EFormOfStudy.MixedTime) &&
                //    rpd.EducationalWorks[EFormOfStudy.MixedTime].Table == null && 
                //    m_regexMixedTimeTable.IsMatch(par.Text)) {
                //    rpd.EducationalWorks[EFormOfStudy.MixedTime].Table = table;
                //    result = true;
                //    break;
                //}
                //if (rpd.EducationalWorks.ContainsKey(EFormOfStudy.PartTime) &&
                //    rpd.EducationalWorks[EFormOfStudy.PartTime].Table == null && 
                //    m_regexPartTimeTable.IsMatch(par.Text)) {
                //    rpd.EducationalWorks[EFormOfStudy.PartTime].Table = table;
                //    result = true;
                //    break;
                //}
            }

            return result;
        }
    }
}
