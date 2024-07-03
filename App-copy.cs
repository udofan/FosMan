using FastMember;
using Microsoft.VisualBasic;
using Microsoft.Web.WebView2.Core;
using System;
using System.CodeDom;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.DirectoryServices.ActiveDirectory;
using System.Drawing;
using System.Drawing.Design;
using System.Globalization;
using System.Linq;
using System.Numerics;
using System.Printing;
using System.Reflection;
using System.Security.Cryptography.Xml;
using System.Security.Policy;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Text.Unicode;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;
using System.Windows;
using System.Windows.Controls;
//using System.Windows.Documents;
using System.Windows.Media.Media3D;
using Xceed.Document.NET;
using Xceed.Pdf.Layout.Table;
using Xceed.Words.NET;
using static FosMan.Enums;
using static System.Resources.ResXFileRef;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

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
        //тест таблицы паспорта ФОСа
        static Regex m_regexFosPassportHeaderTopic = new(@"(модул|раздел|тем)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexFosPassportHeaderIndicator = new(@"индикатор", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static Regex m_regexFosPassportHeaderEvalTool = new(@"оценоч", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        static CompetenceMatrix m_competenceMatrix = null;
        static ConcurrentDictionary<string, Curriculum> m_curriculumDic = [];
        static ConcurrentDictionary<string, Rpd> m_rpdDic = [];
        static ConcurrentDictionary<string, Fos> m_fosDic = [];
        static ConcurrentDictionary<string, CurriculumGroup> m_curriculumGroupDic = [];

        //static Dictionary<string, Department> m_departments = [];
        static Config m_config = new();
        static JsonSerializerOptions m_jsonOptions = new() {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic),
            Converters = {
                new JsonStringEnumConverter()
            }
        };
        static Store m_store = new();

        public delegate void RpdAddEventHandler(Rpd rpd);
        public static event RpdAddEventHandler RpdAdd = null;

        public static CompetenceMatrix CompetenceMatrix { get => m_competenceMatrix; }

        public static ConcurrentDictionary<string, Rpd> RpdList { get => m_rpdDic; }
        public static ConcurrentDictionary<string, Fos> FosList { get => m_fosDic; }

        /// <summary>
        /// Конфигурация
        /// </summary>
        public static Config Config { get => m_config; }

        /// <summary>
        /// Учебные планы в сторе
        /// </summary>
        public static ConcurrentDictionary<string, Curriculum> Curricula { get => m_curriculumDic; }
        /// <summary>
        /// Группы УП
        /// </summary>
        public static ConcurrentDictionary<string, CurriculumGroup> CurriculumGroups { get => m_curriculumGroupDic; }

        /// <summary>
        /// Стор данных
        /// </summary>
        public static Store Store { get => m_store; }

        static App() {
            //m_store = new();
            LoadStore();
        }

        /// <summary>
        /// Список кафедр
        /// </summary>
        //public static Dictionary<string, Department> Departments { get => m_departments; }

        public static bool HasCurriculumFile(string fileName) => m_curriculumDic.ContainsKey(fileName);

        public static bool HasRpdFile(string fileName) => m_rpdDic.ContainsKey(fileName);
        public static bool HasFosFile(string fileName) => m_fosDic.ContainsKey(fileName);

        /// <summary>
        /// Добавить УП в общий стор
        /// </summary>
        /// <param name="curriculum"></param>
        /// <returns></returns>
        static public bool AddCurriculum(Curriculum curriculum, bool addToStore = true) {
            var result = false;
            if (m_curriculumDic.TryAdd(curriculum.SourceFileName, curriculum)) {
                m_curriculumGroupDic.AddOrUpdate(curriculum.GroupKey, key => new(curriculum), (key, value) => {
                    value.AddCurriculum(curriculum);
                    return value;
                });

                //});.TryGetValue(curriculum.GroupKey, out var group)) {
                //    group = new() {
                //        Department = curriculum.Department,
                //        DirectionCode = curriculum.DirectionCode,
                //        DirectionName = curriculum.DirectionName,
                //        Profile = curriculum.Profile,
                //        FSES = curriculum.FSES
                //    };
                //    m_curriculumGroupDic[curriculum.GroupKey] = group;
                //}
                //group.AddCurriculum(curriculum);
                CheckDisciplines();
                result = true;

                if (addToStore) {
                    m_store.CurriculaDic.TryAdd(curriculum.SourceFileName, curriculum);
                }
            }

            return result;
        }

        /// <summary>
        /// Добавить в список новый РПД
        /// </summary>
        /// <param name="fos"></param>
        /// <returns></returns>
        static public bool AddRpd(Rpd rpd, bool addToStore = true) {
            if (rpd == null) return false;

            m_rpdDic[rpd.SourceFileName] = rpd;
            RpdAdd?.Invoke(rpd);

            if (addToStore) {
                m_store.RpdDic.TryAdd(rpd.SourceFileName, rpd);
            }

            return true;
        }

        /// <summary>
        /// Добавить в список новый ФОС
        /// </summary>
        /// <param name="fos"></param>
        /// <returns></returns>
        static public bool AddFos(Fos fos, bool addToStore = true) {
            if (fos == null) return false;

            m_fosDic[fos.SourceFileName] = fos;

            if (addToStore) {
                m_store.FosDic.TryAdd(fos.SourceFileName, fos);
            }

            return true;
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

        /// <summary>
        /// Extension для ячейки: установить текст первому абзацу
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="joinParText"></param>
        /// <returns></returns>
        static public void SetText(this Cell cell, string text = "", bool removeAllParagraphsExceptFirstOne = true) {
            if (removeAllParagraphsExceptFirstOne) {
                cell.Clear();
            }

            var par = cell.Paragraphs.FirstOrDefault();

            if (par != null) {
                var replaceTextOptions = new FunctionReplaceTextOptions() {
                    FindPattern = @"^.*$",
                    ContainerLocation = ReplaceTextContainer.All,
                    RegexMatchHandler = m => text
                };
                par.ReplaceText(replaceTextOptions);
            }
        }

        /// <summary>
        /// Установить значение в ячейку
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <param name="ifCaseOfZero"></param>
        static public void SetText(this Cell cell, int value, string ifCaseOfZero = "-", bool removeAllParagraphsExceptFirstOne = true) {
            var text = value > 0 ? value.ToString() : ifCaseOfZero;
            SetText(cell, text, removeAllParagraphsExceptFirstOne);
        }

        /// <summary>
        /// Очистка ячейки [МЕТОД НЕ РАБОТАЕТ, пустые параграфы не удаляются]
        /// </summary>
        /// <param name="cell"></param>
        static public void Clear(this Cell cell) {
            var replaceTextOptions = new FunctionReplaceTextOptions() {
                FindPattern = @"^.*$",
                ContainerLocation = ReplaceTextContainer.All,
                RegexMatchHandler = m => string.Empty,
                RemoveEmptyParagraph = true,
                TrackChanges = true
            };
            cell.ReplaceText(replaceTextOptions);
            while (cell.Paragraphs.Count > 1) {
                if (!cell.RemoveParagraph(cell.Paragraphs.Last())) {
                    break;
                }
                //cell.Paragraphs.LastOrDefault()?.Remove(false);
            }
        }

        static void AddError(this StringBuilder report, string error) {
            report.Append($"<div style='color: red'>{error}</div>");
        }

        static void AddDiv(this StringBuilder report, string Summary) {
            report.Append($"<div>{Summary}</div>");
        }

        static void AddFileLink(this StringBuilder report, string text, string fileName) {
            //report.Append($"<div>{text} <a href='file://{fileName.Replace(@"\",@"/")}'>{fileName}</a></div>");
            report.Append($"<div>{text} <span onclick='window.chrome.webview.hostObjects.external.jsOpenFile(this.innerText);'" +
                $" style='text-decoration: underline; cursor: hand; color: blue'>{fileName}</span></div>");
        }

        static void AddTocElement(this StringBuilder toc, string element, string anchor, int errorCount) {
            var color = errorCount > 0 ? "red" : "green";
            toc.Append($"<li><a href='#{anchor}'><span style='color:{color};'>{element} (ошибок: {errorCount})</span></a></li>");
        }

        /// <summary>
        /// Добавить в Html-отчет строку в таблицу проверок
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rowNum"></param>
        /// <param name="description"></param>
        /// <param name="result"></param>
        /// <param name="comment"></param>
        static void AddReportTableRow(this StringBuilder table, int rowNum, string description, bool result, string comment) {
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
            (bool result, string msg) ret = (false, "?");

            try {
                rpd.EducationalWorks.TryGetValue(formOfStudy, out EducationalWork eduWork);
                ret = func.Invoke(eduWork);
                if (!ret.result) errorCount++;
            }
            catch (Exception ex) {
                ret.result = false;
                ret.msg = $"<b>{ex.Message}</b><br />{ex.StackTrace}";
            }

            repTable.AddReportTableRow(pos, description, ret.result, ret.msg);

            return ret.result;
        }

        /// <summary>
        /// Проверка ФОС
        /// </summary>
        /// <param name="rpd"></param>
        /// <param name="discipline"></param>
        /// <param name="repTable"></param>
        static bool ApplyFosCheck(Fos fos, Rpd rpd, StringBuilder repTable,
                                  ref int pos, ref int errorCount, string description,
                                  Func<Fos, Rpd, (bool result, string msg)> func) {
            pos++;
            (bool result, string msg) ret = (false, "?");

            try {
                ret = func.Invoke(fos, rpd);
                if (!ret.result) errorCount++;
            }
            catch (Exception ex) {
                ret.result = false;
                ret.msg = $"<b>{ex.Message}</b><br />{ex.StackTrace}";
            }

            repTable.AddReportTableRow(pos, description, ret.result, ret.msg);

            return ret.result;
        }

        /// <summary>
        /// Проверить список РПД
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
                rpd.ExtraErrors = new();
                var errorCount = 0;

                //var discipName = rpd.DisciplineName;
                //if (string.IsNullOrEmpty(discipName)) discipName = "?";
                rep.Append($"<div id='{anchor}' style='width: 100%;'><h3 style='background-color: lightsteelblue'>{rpd.DisciplineName ?? "?"}</h3>");
                rep.Append("<div style='padding-left: 30px'>");
                rep.AddFileLink($"Файл РПД:", rpd.SourceFileName);
                //ошибки, выявленные при загрузке
                if (!rpd.Errors.IsEmpty) {
                    errorCount += rpd.Errors.Count;
                    rep.Append("<p />");
                    rep.AddDiv($"<div style='color: red'>Ошибки, выявленные при загрузке РПД ({rpd.Errors.Count} шт.):</div>");
                    rpd.Errors.Items.ForEach(e => rep.AddError(e.ToString()));
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
                        rep.AddFileLink($"Файл УП:", curriculum.Value.SourceFileName);
                        var discipline = curriculum.Value.FindDiscipline(rpd.DisciplineName);
                        var checkPos = 0;
                        if (discipline != null) {
                            rep.AddDiv($"Проверка по УП:");
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
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Сводная таблица учебных работ: итоговое время", (eduWork) => {
                                    var result = discipline.EducationalWork.TotalHours == eduWork.TotalHours;
                                    var msg = result ? "" : $"Итоговое время [{discipline.EducationalWork.TotalHours}] не соответствует УП (д.б. {eduWork.TotalHours}).";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Сводная таблица учебных работ: время контактной работы", (eduWork) => {
                                    var result = (discipline.EducationalWork.ContactWorkHours ?? 0) == (eduWork.ContactWorkHours ?? 0);
                                    var msg = result ? "" : $"Время контактной работы [{eduWork.ContactWorkHours}] не соответствует значению из УП [{discipline.EducationalWork.ContactWorkHours}].";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Сводная таблица учебных работ: время контроля", (eduWork) => {
                                    var result = (discipline.EducationalWork.ControlHours ?? 0) == (eduWork.ControlHours ?? 0);
                                    var msg = result ? "" : $"Время контроля [{eduWork.ControlHours}] не соответствует значению из УП [{discipline.EducationalWork.ControlHours}].";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Сводная таблица учебных работ: время самостоятельной работы", (eduWork) => {
                                    var result = (discipline.EducationalWork.SelfStudyHours ?? 0) == (eduWork.SelfStudyHours ?? 0);
                                    var msg = result ? "" : $"Время самостоятельных работ [{eduWork.SelfStudyHours}] не соответствует значению из УП [{discipline.EducationalWork.SelfStudyHours}].";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Сводная таблица учебных работ: время практических работ", (eduWork) => {
                                    var result = (discipline.EducationalWork.PracticalHours ?? 0) == (eduWork.PracticalHours ?? 0);
                                    var msg = result ? "" : $"Время практических работ [{eduWork.PracticalHours}] не соответствует значению из УП [{discipline.EducationalWork.PracticalHours}].";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Сводная таблица учебных работ: время лекций", (eduWork) => {
                                    var result = (discipline.EducationalWork.LectureHours ?? 0) == (eduWork.LectureHours ?? 0);
                                    var msg = result ? "" : $"Время лекций [{eduWork.LectureHours}] не соответствует из УП [{discipline.EducationalWork.LectureHours}].";
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
                                if (result && !rpd.CompetenceMatrix.Errors.IsEmpty) {
                                    msg = "В матрице обнаружены ошибки:<br />";
                                    msg += string.Join("<br />", rpd.CompetenceMatrix.Errors.Items.Select(e => $"{e}"));
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
                            //проверка таблицы учебных работ для текущей формы обучения curriculum.Key
                            ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount,
                                $"Проверка времени в таблице содержания ({curriculum.Key.GetDescription()})",
                                eduWork => {
                                    if (eduWork.Table == null) {
                                        return (false, "Таблица не найдена");
                                    }
                                    else {
                                        var msg = "";
                                        try {
                                            var result = true;
                                            //var rowCount = eduWork.Table.RowCount - 4;
                                            //проверка по столбцам и по рядам
                                            var cells = new int[eduWork.Table.RowCount, eduWork.Table.ColumnCount - 2];
                                            var startRow = 3; // eduWork.Table.RowCount - ;
                                            var startCol = -1;
                                            //первый ряд - ряд без объединений ячеек
                                            var firstNumCol = -1;
                                            var lastNumCol = -1;
                                            while (startRow < eduWork.Table.RowCount) {
                                                firstNumCol = -1;
                                                lastNumCol = -1;
                                                var hasNumCells = false;
                                                for (var col = 0; col < eduWork.Table.Rows[startRow].Cells.Count; col++) {
                                                    var cell = eduWork.Table.Rows[startRow].Cells[col];
                                                    //if (cell.GridSpan > 0) {
                                                    //    startRow++;
                                                    //    continue;
                                                    //}
                                                    var text = cell.GetText();
                                                    if (int.TryParse(text, out _)) {
                                                        hasNumCells = true;
                                                        if (firstNumCol < 0) firstNumCol = col;
                                                        if (col > lastNumCol) lastNumCol = col;
                                                        if (col > 0) {
                                                            if (startCol < 0) startCol = col;
                                                        }
                                                    }
                                                }
                                                if (hasNumCells) {
                                                    break;
                                                }
                                                startRow++;
                                            }
                                            //тестируем на кол-во ячеек с числовыми значениями: должно быть 4 или 5 (когда есть подитог по КР)
                                            var hasSubTotal = lastNumCol - firstNumCol + 1 == 5;

                                            //var startCol = eduWork.Table.ColumnCount - 6;

                                            for (int row = startRow; row < cells.GetLength(0) - 1; row++) {
                                                for (int col = startCol; col < cells.GetLength(1); col++) {
                                                    var cellValue = eduWork.Table.Rows[row].Cells[col].GetText();
                                                    if (!string.IsNullOrEmpty(cellValue) && int.TryParse(cellValue, out var intVal)) {
                                                        cells[row, col] = intVal;
                                                    }
                                                }
                                                //проверка строки
                                                if (row < cells.GetLength(0) - 2) { //последние 2 строки пропускаем
                                                    //итог контактной работы
                                                    if (hasSubTotal) {
                                                        var subTotal = cells[row, startCol + 2] + cells[row, startCol + 3];
                                                        if (cells[row, startCol + 1] != subTotal) {
                                                            if (msg.Length > 0) msg += "<br />";
                                                            msg += $"Ряд <b>{eduWork.Table.Rows[row].Cells[startCol - 1].GetText()}</b>: значение [{cells[row, startCol + 1]}] в колонке [Контактная работа.Всего часов] не совпадает с суммой [{subTotal}]";
                                                        }
                                                    }
                                                    //обшее кол-во часов
                                                    var colDelta = hasSubTotal ? 1 : 0;
                                                    var total = cells[row, startCol + 1 + colDelta] + cells[row, startCol + 2 + colDelta] + cells[row, startCol + 3 + colDelta];
                                                    if (cells[row, startCol] != total) {
                                                        if (msg.Length > 0) msg += "<br />";
                                                        msg += $"Ряд <b>{eduWork.Table.Rows[row].Cells[startCol - 1].GetText()}</b>: значение [{cells[row, startCol]}] в колонке [Общее к-во часов] не совпадает с суммой [{total}]";
                                                    }
                                                }
                                            }
                                            //подсчет нижнего итогового ряда (в нем могут быть объединения ячеек)
                                            var total1stCol = 0;
                                            var physColIdx = 0;
                                            while (physColIdx < startCol) {
                                                physColIdx += eduWork.Table.Rows[eduWork.Table.RowCount - 1].Cells[total1stCol].GridSpan;
                                                physColIdx++;
                                                total1stCol++;
                                            }
                                            var matrixCol = startCol - 1;
                                            for (var c = total1stCol; c < eduWork.Table.Rows[eduWork.Table.RowCount - 1].Cells.Count; c++) {
                                                matrixCol++;
                                                var text = eduWork.Table.Rows[eduWork.Table.RowCount - 1].Cells[c].GetText();
                                                if (int.TryParse(text, out var intVal)) {
                                                    cells[eduWork.Table.RowCount - 1, matrixCol] = intVal;
                                                }
                                            }

                                            //проверка итоговой строки
                                            Func<int, int> colSum = colIdx => {
                                                int sum = 0;
                                                for (int row = startRow; row < cells.GetLength(0) - 1; row++) {
                                                    sum += cells[row, colIdx];
                                                }
                                                return sum;
                                            };

                                            var lastRow = eduWork.Table.RowCount - 1;
                                            for (var col = startCol; col < cells.GetLength(1); col++) {
                                                if (cells[lastRow, col] != colSum(col)) {
                                                    var header = col.ToString();
                                                    if (msg.Length > 0) msg += "<br />";
                                                    msg += $"Колонка <b>{header}</b>: итоговая сумма [{colSum(col)}] не совпадает с необходимой [{cells[lastRow, col]}]";
                                                }
                                            }
                                            //проверяем по сводной таблице
                                            if (eduWork.TotalHours != cells[lastRow, startCol]) {
                                                if (msg.Length > 0) msg += "<br />";
                                                msg += $"Колонка <b>(#{startCol}) [Общее время]</b>: значение [{cells[lastRow, startCol]}] не совпадает со значением из сводной таблицы - {eduWork.TotalHours}";
                                            }
                                            var colIdx = startCol + 1 + (hasSubTotal ? 1 : 0);
                                            if (eduWork.LectureHours != cells[lastRow, colIdx]) {
                                                if (msg.Length > 0) msg += "<br />";
                                                msg += $"Колонка <b>(#{colIdx}) [Контактная работа.Лекции]</b>: значение [{cells[lastRow, colIdx]}] не совпадает со значением из сводной таблицы - {eduWork.LectureHours}";
                                            }
                                            colIdx++;
                                            if (eduWork.PracticalHours != cells[lastRow, colIdx]) {
                                                if (msg.Length > 0) msg += "<br />";
                                                msg += $"Колонка <b>(#{colIdx}) [Контактная работа.Практика]</b>: значение [{cells[lastRow, colIdx]}] не совпадает со значением из сводной таблицы - {eduWork.PracticalHours}";
                                            }
                                            colIdx++;
                                            if (eduWork.SelfStudyHours != cells[lastRow, colIdx]) {
                                                if (msg.Length > 0) msg += "<br />";
                                                msg += $"Колонка <b>(#{colIdx}) [Самостоятельная работа]</b>: значение [{cells[lastRow, colIdx]}] не совпадает со значением из сводной таблицы - {eduWork.SelfStudyHours}";
                                            }
                                        }
                                        catch (Exception ex2) {
                                            msg = $"<b>{ex2.Message}</b><br />{ex2.StackTrace}";
                                        }

                                        return (string.IsNullOrEmpty(msg), msg);
                                    }
                                    //var result = discipline.EducationalWork.TotalHours == eduWork.TotalHours;
                                    //var msg = result ? "" : $"Итоговое время [{discipline.EducationalWork.TotalHours}] не соответствует УП (д.б. {eduWork.TotalHours}).";
                                });

                            //проверка тем в таблицах учебных работ по формам обучения
                            ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, $"Проверка списка тем в таблице содержания ({curriculum.Key.GetDescription()})", (eduWork) => {
                                var msg = "";
                                foreach (var item in rpd.EducationalWorks) {
                                    if (item.Value != eduWork) {
                                        var idx = 0;
                                        foreach (var module in eduWork.Modules) {
                                            idx++;
                                            var normalizedTopic = NormalizeText(module.Topic);
                                            if (item.Value.Modules.FirstOrDefault(m => NormalizeText(m.Topic).Equals(normalizedTopic)) == null) {
                                                if (msg.Length > 0) msg += "<br />";
                                                msg += $"Тема <b>{module.Topic}</b> не найдена в таблице содержания формы обучения [{item.Key.GetDescription()}]";
                                            }
                                        }
                                    }
                                    else {
                                        var uniqueEvalTools = new HashSet<EEvaluationTool>();
                                        foreach (var module in eduWork.Modules) {
                                            if (string.IsNullOrEmpty(module.Topic)) {
                                                if (msg.Length > 0) msg += "<br />";
                                                msg += $"Тема #{idx} не указана [{item.Key.GetDescription()}]";
                                            }
                                            else {
                                                if (module.EvaluationTools == null || module.EvaluationTools.Count == 0) {
                                                    if (msg.Length > 0) msg += "<br />";
                                                    msg += $"Тема <b>{module.Topic}</b>: не указано оценочное средство [{item.Key.GetDescription()}]";
                                                }
                                                else {
                                                    foreach (var evalTool in module.EvaluationTools) {
                                                        uniqueEvalTools.Add(evalTool);
                                                    }
                                                }
                                                if (module.CompetenceResultCodes == null || module.CompetenceResultCodes.Count == 0) {
                                                    if (msg.Length > 0) msg += "<br />";
                                                    msg += $"Тема <b>{module.Topic}</b>: не указаны коды результатов индикаторов компетенции [{item.Key.GetDescription()}]";
                                                }
                                                else { //проверка компетенций по матрице
                                                    var results = rpd.CompetenceMatrix.GetAllResultCodes();
                                                    foreach (var code in module.CompetenceResultCodes) {
                                                        if (!results.Contains(code)) {
                                                            if (msg.Length > 0) msg += "<br />";
                                                            msg += $"Тема <b>{module.Topic}</b>: указанный код индикатора компетенции [{code}] не найден в матрице";
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (uniqueEvalTools.Count < 3) {
                                            if (msg.Length > 0) msg += "<br />";
                                            msg += $"Слишком маленький список применяемых оценочных средств ({uniqueEvalTools.Count} шт.): " +
                                                   $"{string.Join(", ", uniqueEvalTools.Select(t => t.GetDescription()))}";
                                        }
                                    }
                                }
                                return (string.IsNullOrEmpty(msg), msg);
                            });

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
        /// Проверить список ФОС
        /// </summary>
        public static void CheckFos(List<Fos> fosList, out string htmlReport) {
            htmlReport = string.Empty;

            StringBuilder html = new("<html><body><h2>Отчёт по проверке ФОС</h2>");
            StringBuilder toc = new("<div><ul>");
            StringBuilder rep = new("<div>");
            var idx = 0;
            foreach (var fos in fosList) {
                var anchor = $"fos{idx}";
                idx++;
                fos.ExtraErrors = [];
                var errorCount = 0;

                //var discipName = rpd.DisciplineName;
                //if (string.IsNullOrEmpty(discipName)) discipName = "?";
                rep.Append($"<div id='{anchor}' style='width: 100%;'><h3 style='background-color: lightsteelblue'>{fos.DisciplineName ?? "?"}</h3>");
                rep.Append("<div style='padding-left: 30px'>");
                rep.AddFileLink($"Файл ФОС:", fos.SourceFileName);
                //ошибки, выявленные при загрузке
                if (fos.Errors.Any()) {
                    errorCount += fos.Errors.Count;
                    rep.Append("<p />");
                    rep.AddDiv($"<div style='color: red'>Ошибки, выявленные при загрузке ({fos.Errors.Count} шт.):</div>");
                    fos.Errors.ForEach(e => rep.AddError(e));
                }
                //для проверки матрицы по УП ищем Учебные планы и берем первый
                CurriculumDiscipline discipline = null;
                var curriculum = FindCurricula(fos)?.FirstOrDefault().Value;
                if (curriculum != null) {
                    rep.AddFileLink($"Файл УП", curriculum.SourceFileName);
                    discipline = curriculum.FindDiscipline(fos.DisciplineName);
                    if (discipline == null) {
                        errorCount++;
                        rep.AddError($"Не удалось найти дисциплину в учебном плане <b>{curriculum.SourceFileName}</b>.");
                    }
                }
                else {
                    errorCount++;
                    rep.AddError("Не удалось найти учебные планы. Добавление УП осуществляется на вкладке <b>\"Учебные планы\"</b>.");
                    rep.AddError("Проверка матрицы компетенций по УП невозможна.");
                }

                //для проверки нам будет нужен РПД
                var rpd = FindRpd(fos);
                if (rpd != null) {
                    //rep.Append("<p />");
                    rep.AddFileLink("Файл РПД:", rpd.SourceFileName);
                    //rep.AddDiv($"Файл РПД: <b>{rpd.SourceFileName}</b>");
                    rep.Append("<p />");
                    rep.AddDiv("<b>Проверки по РПД:</b>");
                    var tdStyle = " style='border: 1px solid;'";
                    var table = new StringBuilder(@"<table style='border: 1px solid'><tr style='font-weight: bold; background-color: lightgray'>");
                    table.Append($"<th {tdStyle}>№ п/п</th><th {tdStyle}>Проверка</th><th {tdStyle}>Результат</th><th {tdStyle}>Комментарий</th>");
                    table.Append("</tr>");
                    int checkPos = 0;
                    //проверка осн. св-в
                    ApplyFosCheck(fos, rpd, table, ref checkPos, ref errorCount, "Кафедра", (fos, rpd) => {
                        var result = string.Compare(fos.Department, rpd.Department, true) == 0;
                        var msg = result ? "" : $"Обнаружено различие в кафедре [ФОС: <b>{fos.Department}</b>, РПД: <b>{rpd.Department}</b>]";
                        return (result, msg);
                    });
                    ApplyFosCheck(fos, rpd, table, ref checkPos, ref errorCount, "Формы обучения", (fos, rpd) => {
                        var result = !fos.FormsOfStudy.Except(rpd.FormsOfStudy).Any() && !rpd.FormsOfStudy.Except(fos.FormsOfStudy).Any();
                        var msg = result ? "" : $"Обнаружено различие в формах обучения " +
                                               $"[ФОС: <b>{string.Join(", ", fos.FormsOfStudy.Select(f => f.GetDescription()))}</b>], " +
                                               $"[РПД: <b>{string.Join(", ", rpd.FormsOfStudy.Select(f => f.GetDescription()))}</b>]";
                        return (result, msg);
                    });
                    ApplyFosCheck(fos, rpd, table, ref checkPos, ref errorCount, "Год", (fos, rpd) => {
                        var result = string.Compare(fos.Year, rpd.Year, true) == 0;
                        var msg = result ? "" : $"Обнаружено различие в годах [ФОС: <b>{fos.Year}</b>, РПД: <b>{rpd.Year}</b>]";
                        return (result, msg);
                    });

                    //проверяем ФОС на наличие матрицы компетенций
                    var checkCompetences = ApplyFosCheck(fos, rpd, table, ref checkPos, ref errorCount, "Наличие матрицы компетенций в ФОС", (fos, rpd) => {
                        var result = fos.CompetenceMatrix?.IsLoaded ?? false;
                        var msg = result ? "" : $"В ФОС не обнаружена матрица компетенций.";
                        if (result && !fos.CompetenceMatrix.Errors.IsEmpty) {
                            msg = "В матрице обнаружены ошибки:<br />";
                            msg += string.Join("<br />", rpd.CompetenceMatrix.Errors.Items.Select(e => $"{e}"));
                            result = false;
                        }
                        if (!result) {
                            if (msg.Length > 0) msg += "<br />";
                            msg += "Проверка компетенций невозможна";
                        }
                        return (result, msg);
                    });

                    if (checkCompetences) {
                        ApplyFosCheck(fos, rpd, table, ref checkPos, ref errorCount, "Проверка матрицы компетенций по РПД", (fos, rpd) => {
                            var achiCodeList = fos.CompetenceMatrix.GetAllAchievementCodes();
                            var summary = "";
                            var matrixError = false;
                            foreach (var code in rpd.CompetenceMatrix.GetAllAchievementCodes()) {
                                var elem = "";
                                if (achiCodeList.Contains(code)) {
                                    elem = $"<span style='color: green; font-weight: bold'>{code}</span>";
                                }
                                else {
                                    elem = $"<span style='color: red; font-weight: bold'>{code}</span>";
                                    matrixError = true;
                                }

                                if (summary.Length > 0) summary += "; ";
                                summary += elem;
                            }
                            foreach (var missedCode in achiCodeList.Except(rpd.CompetenceMatrix.GetAllAchievementCodes())) {
                                matrixError = true;
                                if (summary.Length > 0) summary += "; ";
                                summary += $"<span style='color: red; font-decoration: italic'>{missedCode}??</span>"; ;
                            }
                            var msg = matrixError ? $"Выявлено несоответствие компетенций: {summary}" : "";
                            return (!matrixError, msg);
                        });

                        //проверка матрицы по УП
                        if (discipline != null) {
                            rep.AddDiv($"Проверка матрицы компетенций по УП:");

                            //проверка компетенций
                            checkCompetences &= ApplyFosCheck(fos, rpd, table, ref checkPos, ref errorCount, "Наличие матрицы компетенций в УП", (fos, rpd) => {
                                var result = discipline.CompetenceList?.Any() ?? false;
                                var msg = result ? "" : $"Не удалось определить матрицу компетенций в УП.";
                                return (result, msg);
                            });

                            if (checkCompetences) {
                                ApplyFosCheck(fos, rpd, table, ref checkPos, ref errorCount, "Проверка компетенций по УП", (fos, rpd) => {
                                    var achiCodeList = fos.CompetenceMatrix.GetAllAchievementCodes();
                                    var summary = "";
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

                                        if (summary.Length > 0) summary += "; ";
                                        summary += elem;
                                    }
                                    foreach (var missedCode in achiCodeList.Except(discipline.CompetenceList)) {
                                        matrixError = true;
                                        if (summary.Length > 0) summary += "; ";
                                        summary += $"<span style='color: red; font-decoration: italic'>{missedCode}??</span>"; ;
                                    }
                                    var msg = matrixError ? $"Выявлено несоответствие компетенций: {summary}" : "";
                                    return (!matrixError, msg);
                                });
                                ApplyFosCheck(fos, rpd, table, ref checkPos, ref errorCount, "Проверка этапов формирований компетенций по УП", (fos, rpd) => {
                                    var msg = "";
                                    foreach (var item in fos.CompetenceMatrix.Items) {
                                        if (item.Semester < discipline.StartSemesterIdx + 1 || item.Semester > discipline.LastSemesterIdx + 1) {
                                            if (msg.Length > 0) msg += "<br />";
                                            msg += $"Компетенция <b>{item.Code}</b>: семестр имеет значение [{item.Semester}], " +
                                                   $"а должен лежать в пределах [{discipline.StartSemesterIdx + 1};{discipline.LastSemesterIdx + 1}]";
                                        }
                                    }
                                    return (string.IsNullOrEmpty(msg), msg);
                                });
                            }
                        }
                    }
                    //проверка паспорта
                    ApplyFosCheck(fos, rpd, table, ref checkPos, ref errorCount, "Проверка паспорта: темы", (fos, rpd) => {
                        var msg = "";
                        //список тем по идее для всех форм обучения одинаковый, поэтому смело берем темы очной формы
                        var rpdModules = rpd.EducationalWorks[EFormOfStudy.FullTime].Modules;
                        if (rpdModules != null) {
                            foreach (var module in fos.Passport) {
                                var topic = NormalizeText(module.Topic);
                                if (rpdModules.FirstOrDefault(m => NormalizeText(m.Topic).Equals(topic)) == null) {
                                    if (msg.Length > 0) msg += "<br />";
                                    msg += $"Тема <b>{module.Topic}</b> не найдена в РПД";
                                }
                            }
                        }
                        else {
                            msg = $"В РПД нет данных о содержании дисциплины формы обучения [{EFormOfStudy.FullTime.GetDescription()}]";
                        }
                        return (string.IsNullOrEmpty(msg), msg);
                    });
                    //todo
                    //из РПД вынимать таблицу тем (с учетом объединения ячеек)
                    //сравнивать ее с паспортом ФОС

                    table.Append("</table>");
                    rep.Append(table);
                }
                else {
                    errorCount++;
                    rep.AddError("Не удалось найти родительский РПД. Добавление РПД осуществляется на вкладке <b>\"РПД\"</b>.");
                    rep.AddError("Проверка ФОС по РПД невозможна.");
                }

                if (errorCount == 0) {
                    rep.Append("<p />");
                    rep.AddDiv("<div style='color: green'>Ошибок не обнаружено.</div>");
                }
                rep.Append("</div></div>");

                toc.AddTocElement(fos.DisciplineName ?? "?", anchor, errorCount);
            }
            rep.Append("</div>");
            toc.Append("</ul></div>");
            html.Append(toc).Append(rep).Append("</body></html>");

            htmlReport = html.ToString();
        }

        /// <summary>
        /// Проверка списка УП
        /// </summary>
        /// <param name="curricula"></param>
        /// <param name="htmlReport"></param>
        public static void CheckCurricula(List<Curriculum> curricula, out string htmlReport) {
            htmlReport = string.Empty;

            StringBuilder html = new("<html><body><h2>Отчёт по проверке РПД</h2>");
            StringBuilder toc = new("<div><ul>");
            StringBuilder rep = new("<div>");

            var idx = 0;
            var anchors = new Dictionary<string, string>();
            foreach (var curriculum in curricula) {
                var caption = $"{curriculum.DirectionCode} {curriculum.DirectionName} - {curriculum.Profile}";
                var anchor = $"curriculum{idx}";
                anchors[curriculum.SourceFileName] = anchor;
                idx++;
                //rpd.ExtraErrors = [];
                var errorCount = 0;

                rep.Append($"<div id='{anchor}' style='width: 100%;'><h3 style='background-color: lightsteelblue;'>{caption}</h3>");
                rep.Append("<div style='padding-left: 30px'>");
                rep.AddFileLink($"Файл УП:", curriculum.SourceFileName);
                //ошибки, выявленные при загрузке
                if (curriculum.Errors.Any()) {
                    errorCount += curriculum.Errors.Count;
                    rep.Append("<p />");
                    rep.AddDiv($"<div style='color: red'>Ошибки, выявленные при загрузке УП ({curriculum.Errors.Count} шт.):</div>");
                    curriculum.Errors.ForEach(e => rep.AddError(e));
                }

                if (errorCount == 0) {
                    rep.Append("<p />");
                    rep.AddDiv("<div style='color: green'>Ошибок не обнаружено.</div>");
                }
                rep.Append("</div></div>");

                toc.AddTocElement(caption, anchor, errorCount);
            }

            //проверка дисциплин
            var anchorDisciplines = "disciplines";
            var discCheckCaption = "Проверка дисциплин";
            rep.Append($"<div id='{anchorDisciplines}' style='width: 100%;'><h3 style='background-color: lightgreen;'>{discCheckCaption}</h3>");

            var discErrCount = 0;

            //формируем сквозной список дисциплин
            var disciplines = new ConcurrentDictionary<string, ConcurrentDictionary<string, CurriculumDiscipline>>();
            Parallel.ForEach(curricula, c => {
                foreach (var disc in c.Disciplines.Values) {
                    disciplines.AddOrUpdate(disc.Key, new ConcurrentDictionary<string, CurriculumDiscipline>() { [disc.Curriculum.SourceFileName] = disc }, (key, oldValue) => {
                        oldValue.TryAdd(disc.Curriculum.SourceFileName, disc);
                        return oldValue;
                    });
                }
            });

            var tdStyle = " style='border: 1px solid;'";
            var table = new StringBuilder(@"<table style='border: 1px solid'><tr style='font-weight: bold; background-color: lightgray'>");

            table.Append($"<tr><th {tdStyle}><b>Дисциплина</b></th><th {tdStyle}><b>Входит в УП (количество)</b></th><th {tdStyle}>Проверка компетенций</th></tr>");

            //disciplines = disciplines.OrderBy(x => x.Value.FirstOrDefault().Value.Name);

            idx = 0;
            foreach (var disc in disciplines.OrderBy(x => x.Value.FirstOrDefault().Value.Name)) {
                idx++;
                var hasError = false;

                var competences = disc.Value.Select(x => string.Join(", ", x.Value.CompetenceList.OrderBy(x => x))).ToHashSet();
                hasError = competences.Count > 1;
                var checkStatus = hasError ? "&times;" : "&check;";
                if (competences.Count > 1) {
                    checkStatus = "";
                    var j = 0;
                    foreach (var item in disc.Value) {
                        j++;
                        var curriculum = item.Value.Curriculum;
                        if (checkStatus.Length > 0) checkStatus += "<br />";
                        var tooltip = $"{curriculum.DirectionCode} {curriculum.DirectionName} - {curriculum.Profile}";
                        checkStatus += $"<a href='#{anchors[curriculum.SourceFileName]}' title='{tooltip}'>{j}</a> - {string.Join(", ", item.Value.CompetenceList)}";

                        //if (item.Value.StartSemesterIdx != item.Value.LastSemesterIdx) {
                        //    var tt = 0;
                        //    if (item.Value.LastSemesterIdx - item.Value.StartSemesterIdx > 1) {
                        //        var yy = 0;
                        //    }
                        //}
                    }
                    //checkStatus = string.Join("<br />", competences);
                }

                //table.Append($"<tr {tdStyle}></tr>")
                table.Append($"<tr style='color: {(hasError ? "red" : "green")}'><td {tdStyle}>{idx}. {disc.Value.FirstOrDefault().Value.Name}</td>" +
                                                                               $"<td {tdStyle}>{disc.Value.Count}</td>" +
                                                                               $"<td {tdStyle}>{checkStatus}</td></tr>");
            }

            table.Append("</table");
            rep.Append(table);

            toc.AddTocElement(discCheckCaption, anchorDisciplines, discErrCount);

            rep.Append("</div>");
            toc.Append("</ul></div>");
            html.Append(toc).Append(rep).Append("</body></html>");

            htmlReport = html.ToString();
        }

        /// <summary>
        /// Получение текста ячейки заголовка с учетом объединения ячеек
        /// </summary>
        /// <param name="table"></param>
        /// <param name="headerRowCount"></param>
        /// <param name="colIdx"></param>
        /// <returns></returns>
        static string GetTableHeaderText(Table table, int headerRowCount, int colIdx) {
            //TODO
            var text = "";

            //for (var row = headerRowCount - 1; row >= 0; row--) {
            //    if (row < table.RowCount && colIdx < table.ColumnCount) {
            //        //table.Rows[0].me
            //        var cellText = table.Rows[row].gr.Cells[colIdx].GetText();
            //        if (!string.IsNullOrEmpty(cellText)) {
            //            text = cellText;
            //            break;
            //        }
            //    }
            //}

            return text;
        }

        /// <summary>
        /// Найти учебные планы по формам обучения
        /// </summary>
        /// <param name="directionCode"></param>
        /// <param name="profile"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public static Dictionary<EFormOfStudy, Curriculum> FindCurricula(Rpd rpd) {
            if (!string.IsNullOrEmpty(rpd.DirectionCode) && !string.IsNullOrEmpty(rpd.Profile)) {
                var items = m_curriculumDic.Values.Where(c => string.Compare(c.DirectionCode, rpd.DirectionCode, true) == 0 &&
                                                              string.Compare(c.Profile, rpd.Profile, true) == 0);

                var dic = items.Distinct().ToDictionary(x => x.FormOfStudy, x => x);
                return dic;
            }
            else {
                return null;
            }
        }

        /// <summary>
        /// Найти учебные планы по формам обучения
        /// </summary>
        /// <param name="directionCode"></param>
        /// <param name="profile"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public static Dictionary<EFormOfStudy, Curriculum> FindCurricula(Fos fos) {
            if (!string.IsNullOrEmpty(fos.DirectionCode) && !string.IsNullOrEmpty(fos.Profile)) {
                var items = m_curriculumDic.Values.Where(c => string.Compare(c.DirectionCode, fos.DirectionCode, true) == 0 &&
                                                              string.Compare(c.Profile, fos.Profile, true) == 0);

                var dic = items.Distinct().ToDictionary(x => x.FormOfStudy, x => x);
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
                                    if (CompetenceMatrix.TestTable(table, out var format) && format == ECompetenceMatrixFormat.Rpd) {
                                        RecreateRpdTableOfCompetences(table, disc, false);
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
        /// Поиск РПД по дисциплине
        /// </summary>
        /// <param name="disc"></param>
        /// <returns></returns>
        public static Rpd? FindRpd(CurriculumDiscipline disc) {
            //var name = disc.Name.Replace('ё', 'е').ToLower();

            //var rpd = m_rpdDic.Values.FirstOrDefault(d => d.DisciplineName.ToLower().Replace('ё', 'е').Equals(name));
            //rpd ??= m_rpdDic.Values.FirstOrDefault(d => d.DisciplineName.ToLower().Replace('ё', 'е').StartsWith(name));
            //rpd ??= m_rpdDic.Values.FirstOrDefault(d => name.StartsWith(d.DisciplineName.ToLower().Replace('ё', 'е')));

            var name = NormalizeName(disc.Name);

            var rpd = m_rpdDic.Values.FirstOrDefault(d => NormalizeName(d.DisciplineName).Equals(name));
            rpd ??= m_rpdDic.Values.FirstOrDefault(d => NormalizeName(d.DisciplineName).StartsWith(name));
            rpd ??= m_rpdDic.Values.FirstOrDefault(d => name.StartsWith(NormalizeName(d.DisciplineName)));

            return rpd; //m_rpdDic.Values.FirstOrDefault(r => r.DisciplineName.Equals(disc.Name, StringComparison.CurrentCultureIgnoreCase));
        }

        /// <summary>
        /// Поиск РПД по ФОСу
        /// </summary>
        /// <param name="disc"></param>
        /// <returns></returns>
        public static Rpd? FindRpd(Fos fos) => m_rpdDic.Values.FirstOrDefault(d => d.Key.Equals(fos.Key));

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
        /// <param name="table">таблица</param>
        /// <param name="discipline"></param>
        static bool RecreateRpdTableOfCompetences(Table table, CurriculumDiscipline discipline, bool recreateHeaders) {
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
        /// Установка заголовков для таблицы
        /// </summary>
        /// <param name="table"></param>
        /// <param name="captions"></param>
        /// <param name="bold"></param>
        /// <param name="fontSize"></param>
        /// <param name="lineSpacing"></param>
        /// <param name="alignment"></param>
        /// <returns></returns>
        static bool SetTableHeaders(Table table,
                                    string[] captions,
                                    bool? bold = true,
                                    double? fontSize = 12,
                                    float lineSpacing = 12,
                                    Alignment alignment = Alignment.center,
                                    Xceed.Document.NET.VerticalAlignment verticalAlignment = Xceed.Document.NET.VerticalAlignment.Center) {
            var result = false;
            var format = new Formatting() { Bold = bold, Size = fontSize };

            if (table.RowCount > 0 && (captions?.Any() ?? false)) {
                //if (table.RowCount == 0) {
                //    table.row
                //    var row = new Row();
                //    table.Rows.Add(row);
                //    table.InsertRow();
                //}

                //вставляем второй ряд, который станет заголовком
                var row = table.InsertRow(1);
                table.RemoveRow(0); //а текущий первый ряд удаляем

                var idx = 0;
                foreach (var cell in table.Rows[0].Cells) {
                    if (cell.GridSpan == 0) {  //ячейка должна быть без объединения
                        //cell.Clear();
                        var cellPar = cell.Paragraphs.FirstOrDefault();
                        var firstLine = true;
                        foreach (var line in captions[idx].Split("\r\n")) {
                            //if (!firstLine) cellPar = cellPar.InsertParagraphAfterSelf()
                            cellPar.SetLineSpacing(LineSpacingType.Line, lineSpacing);
                            if (firstLine) {
                                cellPar.InsertText(line, false, formatting: format);
                            }
                            else {
                                cellPar = cell.InsertParagraph(line, false, formatting: format);
                                //var newPar = cellPar.InsertParagraphAfterSelf(line, false, formatting: format);
                                //cellPar = newPar; // cellPar.InsertParagraphAfterSelf(line, false, formatting: format);
                            }
                            cellPar.SetLineSpacing(LineSpacingType.Line, lineSpacing);
                            cellPar.Alignment = alignment;
                            cellPar.IndentationFirstLine = 0.1f;
                            cellPar.LineSpacingAfter = 0.1f;
                            cellPar.LineSpacingBefore = 0.1f;
                            firstLine = false;
                        }
                        cell.VerticalAlignment = verticalAlignment;
                    };
                    idx++;
                    if (idx >= captions.Length) break;
                }

                result = true;
            }

            return result;
        }

        /// <summary>
        /// Воссоздание таблицы компетенций #1 ФОС (п. 2.1)
        /// </summary>
        /// <param name="table"></param>
        /// <param name="discipline"></param>
        /// <returns></returns>
        static bool RecreateFosTableOfCompetences1(Table table, Rpd rpd, bool recreateHeaders, double? fontSize = 12, float lineSpacing = 12) {
            var result = false;

            //очистка таблицы
            if (table.ColumnCount < 3) {
                return false;
            }
            var topRowIdxToRemove = recreateHeaders ? 0 : 1;
            for (var rowIdx = table.RowCount - 1; rowIdx >= topRowIdxToRemove; rowIdx--) {
                table.RemoveRow(rowIdx);
            }

            if (recreateHeaders) {
                //заголовок
                SetTableHeaders(table, [ "Код и наименование\r\nкомпетенций",
                                         "Коды и индикаторы\r\nдостижения компетенций",
                                         "Этапы формирования компетенций (семестр)" ],
                                         fontSize: fontSize,
                                         lineSpacing: lineSpacing);
                //var format = new Formatting() { Bold = true, Size = fontSize };
                //var header0 = table.Rows[0].Cells[0].Paragraphs.FirstOrDefault();
                //header0.InsertText("Код и наименование", false, formatting: format);
                //header0.Alignment = Alignment.center;
                //header0.IndentationFirstLine = 0.1f;
                //var header02 = header0.InsertParagraphAfterSelf("компетенций", false, formatting: format);
                //header02.Alignment = Alignment.center;
                //header02.IndentationFirstLine = 0.1f;

                //var header1 = table.Rows[0].Cells[1].Paragraphs.FirstOrDefault();
                //header1.InsertText("Коды и индикаторы", formatting: new Formatting() { Bold = true });
                //header1.Alignment = Alignment.center;
                //header1.IndentationFirstLine = 0.1f;
                //var header12 = header1.InsertParagraphAfterSelf("достижения компетенций", false, formatting: format);
                //header12.Alignment = Alignment.center;
                //header12.IndentationFirstLine = 0.1f;

                //var header2 = table.Rows[0].Cells[2].Paragraphs.FirstOrDefault();
                //header2.InsertText("Этапы формирования компетенций (семестр)", formatting: format);
                //header2.Alignment = Alignment.center;
                //header2.IndentationFirstLine = 0.1f;
            }

            var discipline = FindDiscipline(rpd);

            //ряды матрицы
            //var appliedIndicatorCount = 0;
            //var items = rpd.CompetenceMatrix.GetItems(discipline.CompetenceList); //отбор строк матрицы по списку индикаторов
            foreach (var item in rpd.CompetenceMatrix.Items) {
                var competenceStartRow = table.RowCount;

                foreach (var achi in item.Achievements) {
                    var resRowIdx = table.RowCount;
                    var newRow = table.InsertRow();

                    if (resRowIdx == competenceStartRow) {
                        //ячейка компетенции
                        var cellCompetence = newRow.Cells[0].Paragraphs.FirstOrDefault();
                        cellCompetence.SetLineSpacing(LineSpacingType.Line, lineSpacing);
                        cellCompetence.Alignment = Alignment.both;
                        cellCompetence.InsertText($"{item.Code}.", formatting: new Formatting() { Bold = true, Size = fontSize });
                        cellCompetence.InsertText($" {item.Title}", formatting: new Formatting() { Bold = false, Size = fontSize });
                        cellCompetence.IndentationFirstLine = 0.1f;
                        //ячейка этапа
                        var cellStages = newRow.Cells[2].Paragraphs.FirstOrDefault();
                        var stages = "";
                        for (var sem = discipline.StartSemesterIdx + 1; sem <= discipline.LastSemesterIdx + 1; sem++) {
                            if (stages.Length > 0) stages += ", ";
                            stages += $"{sem}";
                        }
                        cellStages.InsertText(stages, formatting: new Formatting() { Bold = false, Size = fontSize });
                        cellStages.IndentationFirstLine = 0.1f;
                        cellStages.Alignment = Alignment.center;
                        newRow.Cells[2].VerticalAlignment = Xceed.Document.NET.VerticalAlignment.Center;
                    }

                    //ячейка индикатора
                    var cellIndicator = newRow.Cells[1].Paragraphs.FirstOrDefault();
                    cellIndicator.Alignment = Alignment.both;
                    cellIndicator.SetLineSpacing(LineSpacingType.Line, lineSpacing);
                    cellIndicator.InsertText($"{achi.Code}. {achi.Indicator}", formatting: new Formatting() { Bold = false, Size = fontSize });
                    cellIndicator.IndentationFirstLine = 0.1f;

                    //if (!discipline.CompetenceList.Contains(achi.Code)) { //доп. фильтрация только по нужным индикаторам
                    //    continue;
                    //}
                    //appliedIndicatorCount++;

                    //var achiStartRow = table.RowCount;

                    //foreach (var res in achi.Results) {
                    //    var resRowIdx = table.RowCount;
                    //    var newRow = table.InsertRow();
                    //    if (resRowIdx == competenceStartRow) {
                    //        //ячейка компетенции
                    //        var cellCompetence = newRow.Cells[0].Paragraphs.FirstOrDefault();
                    //        //table.MergeCellsInColumn(0, firstRowIdx, firstRowIdx + rowSpan - 1);
                    //        cellCompetence.InsertText($"{item.Code}.", formatting: new Formatting() { Bold = true });
                    //        cellCompetence.InsertText($" {item.Title}", formatting: new Formatting() { Bold = false });
                    //        cellCompetence.IndentationFirstLine = 0.1f;
                    //    }
                    //    if (resRowIdx == achiStartRow) {
                    //        //ячейка индикатора
                    //        var cellIndicator = newRow.Cells[1].Paragraphs.FirstOrDefault();
                    //        cellIndicator.InsertText($"{achi.Code}. {achi.Indicator}", formatting: new Formatting() { Bold = false });
                    //        cellIndicator.IndentationFirstLine = 0.1f;
                    //    }

                    //    //ячейка результата
                    //    var cellResult = newRow.Cells[2].Paragraphs.FirstOrDefault();
                    //    cellResult.InsertText($"{res.Code}:", formatting: new Formatting() { Bold = false });
                    //    cellResult.IndentationFirstLine = 0.1f;
                    //    var cellResult2 = cellResult.InsertParagraphAfterSelf($"{res.Description}", false, formatting: new Formatting() { Bold = false });
                    //    cellResult2.IndentationFirstLine = 0.1f;
                    //}

                    //if (table.RowCount - 1 > achiStartRow) {
                    //    table.MergeCellsInColumn(1, achiStartRow, table.RowCount - 1);
                    //}
                }
                if (table.RowCount - 1 > competenceStartRow) {
                    table.MergeCellsInColumn(0, competenceStartRow, table.RowCount - 1);
                    table.MergeCellsInColumn(2, competenceStartRow, table.RowCount - 1);
                }
            }
            result = true; // appliedIndicatorCount == discipline.CompetenceList.Count;

            return result;
        }

        /// <summary>
        /// Установить текст в ячейку (пустую изначально)
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        /// <param name="fontSize"></param>
        /// <param name="lineSpacting"></param>
        /// <param name="alignment"></param>
        /// <param name="verticalAlignment"></param>
        public static void SetTextFormatted(this Cell cell, string text,
                                            bool fontBold = false,
                                            double fontSize = 12,
                                            float lineSpacing = 12,
                                            Alignment alignment = Alignment.left,
                                            Xceed.Document.NET.VerticalAlignment verticalAlignment = Xceed.Document.NET.VerticalAlignment.Center) {
            var par = cell.Paragraphs.FirstOrDefault();

            par.SetLineSpacing(LineSpacingType.Line, lineSpacing);
            par.IndentationFirstLine = 0.1f;
            par.Alignment = alignment;
            par.InsertText($"{text}", formatting: new Formatting() { Bold = fontBold, Size = fontSize });
            cell.VerticalAlignment = verticalAlignment;
        }

        /// <summary>
        /// Установить текст в ячейку (пустую изначально)
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="text"></param>
        /// <param name="fontSize"></param>
        /// <param name="lineSpacting"></param>
        /// <param name="alignment"></param>
        /// <param name="verticalAlignment"></param>
        public static void SetTextFormatted(this Cell cell, int value,
                                            string ifCaseOfZero = "-",
                                            bool fontBold = false,
                                            double fontSize = 12,
                                            float lineSpacing = 12,
                                            Alignment alignment = Alignment.left,
                                            Xceed.Document.NET.VerticalAlignment verticalAlignment = Xceed.Document.NET.VerticalAlignment.Center) {
            var par = cell.Paragraphs.FirstOrDefault();

            var text = value > 0 ? value.ToString() : ifCaseOfZero;
            SetTextFormatted(cell, text, fontBold, fontSize, lineSpacing, alignment, verticalAlignment);
        }

        /// <summary>
        /// Воссоздание таблицы компетенций #2 ФОС (п. 2.2)
        /// </summary>
        /// <param name="table">таблица</param>
        /// <param name="discipline"></param>
        static bool RecreateFosTableOfCompetences2(Table table, Rpd rpd, bool recreateHeaders, double? fontSize = 12, float lineSpacing = 12) {
            var result = false;

            //очистка таблицы
            if (table.ColumnCount < 3) {
                return false;
            }
            var topRowIdxToRemove = 1; // recreateHeaders ? 0 : 1;
            for (var rowIdx = table.RowCount - 1; rowIdx >= topRowIdxToRemove; rowIdx--) {
                table.RemoveRow(rowIdx);
            }

            if (recreateHeaders) {
                SetTableHeaders(table, [ "Код\r\nкомпетенции",
                                         "Коды и индикаторы\r\nдостижения\r\nкомпетенций",
                                         "Коды и результаты обучения" ],
                                fontSize: fontSize, lineSpacing: lineSpacing);
                //заголовок
                //var format = new Formatting() { Bold = true, Size = fontSize };
                //var header0 = table.Rows[0].Cells[0].Paragraphs.FirstOrDefault();
                //header0.InsertText("Код", false, formatting: format);
                //header0.Alignment = Alignment.center;
                //header0.IndentationFirstLine = 0.1f;
                //var header02 = header0.InsertParagraphAfterSelf("компетенции", false, formatting: format);
                //header02.Alignment = Alignment.center;
                //header02.IndentationFirstLine = 0.1f;
                //table.Rows[0].Cells[0].VerticalAlignment = Xceed.Document.NET.VerticalAlignment.Center;

                //var header1 = table.Rows[0].Cells[1].Paragraphs.FirstOrDefault();
                //header1.InsertText("Коды и индикаторы", formatting: format);
                //header1.Alignment = Alignment.center;
                //header1.IndentationFirstLine = 0.1f;
                //var header12 = header1.InsertParagraphAfterSelf("достижения", false, formatting: format);
                //header12.Alignment = Alignment.center;
                //header12.IndentationFirstLine = 0.1f;
                //var header13 = header1.InsertParagraphAfterSelf("компетенций", false, formatting: format);
                //header13.Alignment = Alignment.center;
                //header13.IndentationFirstLine = 0.1f;
                //table.Rows[0].Cells[1].VerticalAlignment = Xceed.Document.NET.VerticalAlignment.Center;

                //var header2 = table.Rows[0].Cells[2].Paragraphs.FirstOrDefault();
                //header2.InsertText("Коды и результаты обучения", formatting: format);
                //header2.Alignment = Alignment.center;
                //header2.IndentationFirstLine = 0.1f;
                //table.Rows[0].Cells[2].VerticalAlignment = Xceed.Document.NET.VerticalAlignment.Center;
            }

            var boldFormat = new Formatting() { Bold = true, Size = fontSize };
            var defFormat = new Formatting() { Bold = false, Size = fontSize };

            //ряды матрицы
            foreach (var item in rpd.CompetenceMatrix.Items) {
                var competenceStartRow = table.RowCount;

                foreach (var achi in item.Achievements) {
                    var achiStartRow = table.RowCount;
                    foreach (var res in achi.Results) {
                        var resRowIdx = table.RowCount;
                        var newRow = table.InsertRow();
                        if (resRowIdx == competenceStartRow) {
                            //ячейка компетенции
                            var cellCompetence = newRow.Cells[0].Paragraphs.FirstOrDefault();
                            cellCompetence.SetLineSpacing(LineSpacingType.Line, lineSpacing);
                            cellCompetence.Alignment = Alignment.center;
                            cellCompetence.InsertText($"{item.Code}", formatting: boldFormat);
                            cellCompetence.IndentationFirstLine = 0.1f;
                            newRow.Cells[0].VerticalAlignment = Xceed.Document.NET.VerticalAlignment.Center;
                        }
                        if (resRowIdx == achiStartRow) {
                            //ячейка индикатора
                            var cellIndicator = newRow.Cells[1].Paragraphs.FirstOrDefault();
                            cellIndicator.SetLineSpacing(LineSpacingType.Line, lineSpacing);
                            cellIndicator.Alignment = Alignment.center;
                            cellIndicator.InsertText($"{achi.Code}", formatting: defFormat);
                            cellIndicator.IndentationFirstLine = 0.1f;
                            newRow.Cells[1].VerticalAlignment = Xceed.Document.NET.VerticalAlignment.Center;
                        }

                        //ячейка результата
                        var cellResult = newRow.Cells[2].Paragraphs.FirstOrDefault();
                        cellResult.SetLineSpacing(LineSpacingType.Line, lineSpacing);
                        cellResult.Alignment = Alignment.left;
                        cellResult.InsertText($"{res.Code}:", formatting: defFormat);
                        cellResult.IndentationFirstLine = 0.1f;
                        var cellResult2 = cellResult.InsertParagraphAfterSelf($"{res.Description}", false, formatting: defFormat);
                        cellResult2.SetLineSpacing(LineSpacingType.Line, lineSpacing);
                        cellResult2.Alignment = Alignment.both;
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
            result = true;

            return result;
        }


        /// <summary>
        /// Воссоздание таблицы паспорта ФОСа (п. 3)
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rpd"></param>
        /// <param name="updateHeaders"></param>
        private static bool RecreateFosTableOfPassport(Table table, Rpd rpd, bool recreateHeaders, double? fontSize = 12, float lineSpacing = 12) {
            var result = false;

            if (table.ColumnCount < 4) {
                return false;
            }

            //очистка таблицы
            var topRowIdxToRemove = 1;
            for (var rowIdx = table.RowCount - 1; rowIdx >= topRowIdxToRemove; rowIdx--) {
                table.RemoveRow(rowIdx);
            }

            if (recreateHeaders) {
                SetTableHeaders(table, [ "№ п/п",
                                         "Контролируемые модули, разделы (темы) дисциплины",
                                         "Код контролируемого индикатора достижения компетенции",
                                         "Наименование\r\nоценочного средства" ],
                                fontSize: fontSize, lineSpacing: lineSpacing);
            }

            //цикл по списку модулей
            var idx = 0;
            foreach (var m in rpd.EducationalWorks[EFormOfStudy.FullTime].Modules) {
                idx++;
                var row = table.InsertRow();
                row.Cells[0].SetTextFormatted($"{idx}.", alignment: Alignment.center);
                row.Cells[1].SetTextFormatted($"{m.Topic}");
                //получим список кодов индикаторов по списку результатов
                var indicators = rpd.CompetenceMatrix.GetAchievements(m.CompetenceResultCodes);
                row.Cells[2].SetTextFormatted($"{string.Join(", ", indicators.Select(i => i.Code))}", alignment: Alignment.center);
                row.Cells[3].SetTextFormatted($"{string.Join("\r\n", m.EvaluationTools.Select(t => t.GetDescription()))}", alignment: Alignment.center);
            }

            result = true;
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
        /// Попытка десериализация JSON в объект
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="json"></param>
        /// <returns></returns>
        public static bool TryDeserialize<T>(string json, out T obj) {
            var result = false;
            obj = default;

            try {
                obj = JsonSerializer.Deserialize<T>(json, m_jsonOptions);
                result = true;
            }
            catch (Exception ex) {
                var u = 0;
            }

            return result;
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
            if (!(m_config.FosFixDocPropertyList?.Any() ?? false)) {
                m_config.FosFixDocPropertyList = DocProperty.DefaultProperties;
            }
        }

        /// <summary>
        /// Режим фикса РПД
        /// </summary>
        /// <param name="rpdList"></param>
        internal static void FixRpdFiles(List<Rpd> rpdList, string targetDir, out string htmlReport) {
            var html = new StringBuilder("<html><body><h2>Отчёт по исправлению РПД</h2>");
            var rep = new StringBuilder("");
            DocX templateDocx = null;
            DocxChapters templateChapters = null;
            var successFileCount = 0;
            var failureFileCount = 0;
            var swMain = Stopwatch.StartNew();

            try {
                System.Windows.Forms.Application.UseWaitCursor = true;
                System.Windows.Forms.Application.DoEvents();

                var fixByTemplate = false;
                if (Config.RpdFixByTemplate) {
                    if (File.Exists(Config.RpdFixTemplateFileName)) {
                        templateDocx = DocX.Load(Config.RpdFixTemplateFileName);
                        templateChapters = ScanDocxForChapters(templateDocx);
                        fixByTemplate = true;
                    }
                }

                rep.Append("<div><b>Режим работы:</b></div><ul>");
                if (Config.RpdFixTableOfCompetences) {
                    rep.Append("<li>Исправление таблицы компетенций</li>");
                }
                if (Config.RpdFixTableOfEduWorks) {
                    rep.Append("<li>Исправление таблицы учебных работ</li>");
                }
                if (Config.RpdFixEduWorkTablesFixTime) {
                    rep.Append("<li>Заполнение таблиц учебных работ для форм обучения</li>");
                }
                if (Config.RpdFixSetPrevAndNextDisciplines) {
                    rep.Append("<li>Заполнение списков предыдущих и последующих дисциплин</li>");
                }
                if (Config.RpdFixRemoveColorSelections) {
                    rep.Append("<li>Очистка цветных выделений для служебных областей</li>");
                }
                if (Config.RpdFixEduWorkTablesFullRecreate) {
                    rep.Append("<li>Таблицы содержания дисциплины: полное перестроение</li>");
                }
                if (Config.RpdFixEduWorkTablesFixTime) {
                    rep.Append("<li>Таблицы содержания дисциплины: автоматическое распределение времен по темам</li>");
                }
                if (Config.RpdFixEduWorkTablesFixEvalTools) {
                    rep.Append("<li>Таблицы содержания дисциплины: автоматическое распределение оценочных средств</li>");
                }
                if (Config.RpdFixEduWorkTablesFixCompetenceCodes) {
                    rep.Append("<li>Таблицы содержания дисциплины: автоматическое распределение результатов компетенций</li>");
                }

                List<FindAndReplaceItem> findAndReplaceItems = new();
                if (App.Config.RpdFixFindAndReplace) {
                    findAndReplaceItems = Config.RpdFixFindAndReplaceItems?.Where(i => i.IsChecked).ToList();
                    if (App.Config.RpdFixFindAndReplace && findAndReplaceItems != null && findAndReplaceItems.Any()) {
                        var tdStyle = "style='border: 1px solid;'";
                        rep.Append($"<li><table {tdStyle}><tr><th {tdStyle}><b>Найти</b></th><th {tdStyle}><b>Заменить на</b></th></tr>");
                        foreach (var item in findAndReplaceItems) {
                            rep.Append($"<tr><td {tdStyle}>{item.FindPattern}</td><td {tdStyle}>{item.ReplacePattern}</td></tr>");
                        }
                        rep.Append("</table></li>");
                    }
                }
                rep.Append("</ul>");

                foreach (var rpd in rpdList) {
                    var sw = Stopwatch.StartNew();
                    rep.Append("<p />");
                    if (File.Exists(rpd.SourceFileName)) {
                        rep.AddFileLink($"Исходный файл:", rpd.SourceFileName);

                        try {
                            rep.Append($"<div>Дисциплина: <b>{rpd.DisciplineName}</b></div>");

                            var fixCompetences = Config.RpdFixTableOfCompetences;
                            var discipline = FindDiscipline(rpd);
                            if (discipline == null) {
                                fixCompetences = false;
                                rep.Append($"<div style='color: red'>Не удалось найти дисциплину [{rpd.DisciplineName}] в загруженных учебных планах</div>");
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
                                        rep.Append($"<div style='color: red'>Учебный план для формы обучения [{form.GetDescription()}] не загружен</div>");
                                    }
                                }
                            }
                            //var fixEduWorkTables = Config.RpdFixEduWorkTablesFixTime || Config.RpdFixEduWorkTablesFixEvalTools || Config.RpdFixEduWorkTablesFixCompetenceCodes;
                            EEduWorkFixType fixType = EEduWorkFixType.Undefined;
                            if (Config.RpdFixEduWorkTablesFullRecreate) {
                                fixType = EEduWorkFixType.All;
                            }
                            else {
                                if (Config.RpdFixEduWorkTablesFixTime) fixType |= EEduWorkFixType.Time;
                                if (Config.RpdFixEduWorkTablesFixEvalTools) fixType |= EEduWorkFixType.EvalTools;
                                if (Config.RpdFixEduWorkTablesFixCompetenceCodes) fixType |= EEduWorkFixType.CompetenceResults;
                            }
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

                                EEvaluationTool[] evalTools = null;
                                string[][] studyResults = null;             //здесь будут формироваться значения для таблиц учебных работ по формам обучения

                                foreach (var table in docx.Tables) {
                                    var backup = table.Xml;
                                    if (fixCompetences && CompetenceMatrix.TestTable(table, out var format) && format == ECompetenceMatrixFormat.Rpd) {
                                        if (RecreateRpdTableOfCompetences(table, discipline, false)) {
                                            rep.Append("<div style='color: green'>Таблица компетенций сформирована по матрице компетенций.</div>");
                                        }
                                        else {
                                            rep.Append("<div style='color: red'>Не удалось сформировать обновленную таблицу компетенций.</div>");
                                            table.Xml = backup;
                                        }
                                    }
                                    if (fixEduWorksSummary && !eduSummaryTableIsFixed) {
                                        eduSummaryTableIsFixed |= TestForSummaryTableForEducationalWorks(table, eduWorks, PropertyAccess.Set /*установка значений в таблицу*/);
                                    }
                                    if (fixType != EEduWorkFixType.Undefined) {
                                        //фикс-заполнение таблицы учебных работ для формы обучения
                                        if (TestForEduWorkTable(docx, table, rpd, PropertyAccess.Set, fixType, ref evalTools, ref studyResults,
                                                                Config.RpdFixMaxCompetenceResultsCount,
                                                                Config.RpdFixEduWorkTablesEvalTools1stStageItems,
                                                                Config.RpdFixEduWorkTablesEvalTools2ndStageItems,
                                                                out var formOfStudy)) {
                                            eduWorkTableIsFixed[formOfStudy] = true;
                                        }
                                    }
                                }
                                if (fixEduWorksSummary) {
                                    if (eduSummaryTableIsFixed) {
                                        rep.Append("<div style='color: green'>Сводная таблица учебных работ заполнена по учебным планам.</div>");
                                    }
                                    else {
                                        rep.Append("<div style='color: red'>Не удалось сформировать сводную таблицу учебных работ.</div>");
                                        //table.Xml = backup;
                                    }
                                }
                                if (fixType != EEduWorkFixType.Undefined) {
                                    foreach (var item in eduWorkTableIsFixed) {
                                        if (item.Value) {
                                            rep.Append($"<div style='color: green'>Таблица учебных работ для формы обучения [{item.Key.GetDescription()}] успешно заполнена.</div>");
                                        }
                                        else {
                                            rep.Append($"<div style='color: red'>Таблицу учебных работ для формы обучения [{item.Key.GetDescription()}] не удалось заполнить.</div>");
                                        }
                                    }
                                }

                                //обработка "найти и заменить"
                                var replaceCount = 0;

                                //заполнение предыд. и послед. дисциплин осуществим с помощью поиска и замены
                                List<FindAndReplaceItem> extraReplaceItems = new();

                                if (Config.RpdFixSetPrevAndNextDisciplines) {
                                    extraReplaceItems.Add(new FindAndReplaceItem() {
                                        FindPattern = "{PrevDisciplines}",
                                        ReplacePattern = rpd.PrevDisciplines
                                    });
                                    extraReplaceItems.Add(new FindAndReplaceItem() {
                                        FindPattern = "{NextDisciplines}",
                                        ReplacePattern = rpd.NextDisciplines
                                    });
                                    findAndReplaceItems.AddRange(extraReplaceItems);
                                }

                                if ((findAndReplaceItems?.Any() ?? false) || App.Config.RpdFixRemoveColorSelections) {
                                    foreach (var par in docx.Paragraphs) {
                                        //поиск и замена
                                        if (findAndReplaceItems?.Any() ?? false) {
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
                                        //очистка цветных выделений
                                        if (App.Config.RpdFixRemoveColorSelections) {
                                            par.ShadingPattern(new ShadingPattern() { Fill = Color.Transparent, StyleColor = Color.Transparent }, ShadingType.Paragraph);
                                            par.Highlight(Highlight.none);
                                            //par.lis
                                            //docx.Lists.ForEach(l => l.sele)
                                        }
                                    }
                                    rep.Append($"<div>Осуществлено замен в тексте: {replaceCount}</div>");
                                }
                                extraReplaceItems?.ForEach(x => findAndReplaceItems.Remove(x)); //убираем, т.к. это было нужно только для текущего РПД

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
                                rep.Append($"<div>Время работы: {sw.Elapsed}</div>");
                                rep.AddFileLink($"Итоговый файл:", newFileName);
                            }
                            successFileCount++;
                        }
                        catch (Exception e) {
                            rep.AddError($"При обработке файла возникла ошибка: {e.Message}");
                            rep.AddError($"Стек: {e.StackTrace}");
                            failureFileCount++;
                        }
                    }
                    else {
                        rep.Append($"<div style='color:red'>Файл <b>{rpd.SourceFileName}</b> не найден</div>");
                        failureFileCount++;
                    }
                }
            }
            catch (Exception ex) {
                rep.Append($"<div>{ex.Message}</div>");
                rep.Append($"<div>{ex.StackTrace}</div>");
            }
            finally {
                templateDocx?.Dispose();
                html.AddDiv($"Исправлено файлов: {successFileCount}");
                html.AddDiv($"Сбойных файлов: {failureFileCount}");
                html.AddDiv($"Время работы: {swMain.Elapsed}");
                html.Append("<p />");
                html.Append(rep);
                html.Append("</body></html>");
                System.Windows.Forms.Application.UseWaitCursor = false;
            }

            htmlReport = html.ToString();
        }

        /// <summary>
        /// Режим фикса ФОС
        /// </summary>
        /// <param name="fosList"></param>
        internal static void FixFosFiles(List<Fos> fosList, string targetDir, out string htmlReport) {
            var html = new StringBuilder("<html><body><h2>Отчёт по исправлению ФОС</h2>");
            DocX templateDocx = null;
            DocxChapters templateChapters = null;

            try {
                html.Append("<div><b>Режим работы:</b></div><ul>");
                if (Config.FosFixCompetenceTable1) {
                    html.Append("<li>Исправление таблицы компетенций #1 (п.2.1)</li>");
                }
                if (Config.FosFixCompetenceTable2) {
                    html.Append("<li>Исправление таблицы компетенций #2 (п.2.2)</li>");
                }
                if (Config.FosFixPassportTable) {
                    html.Append("<li>Исправление таблицы паспорта</li>");
                }
                if (Config.FosFixResetSelection) {
                    html.Append("<li>Очистка цветных выделений для служебных областей</li>");
                }
                var findAndReplaceItems = Config.FosFixFindAndReplaceItems?.Where(i => i.IsChecked).ToList();
                if (findAndReplaceItems != null && findAndReplaceItems.Any()) {
                    var tdStyle = "style='border: 1px solid;'";
                    html.Append($"<li><table {tdStyle}><tr><th {tdStyle}><b>Найти</b></th><th {tdStyle}><b>Заменить на</b></th></tr>");
                    foreach (var item in findAndReplaceItems) {
                        html.Append($"<tr><td {tdStyle}>{item.FindPattern}</td><td {tdStyle}>{item.ReplacePattern}</td></tr>");
                    }
                    html.Append("</table></li>");
                }
                html.Append("</ul>");

                foreach (var fos in fosList) {
                    html.Append("<p />");
                    if (File.Exists(fos.SourceFileName)) {
                        html.AddFileLink("Исходный файл:", fos.SourceFileName);
                        //html.Append($"<div>Исходный файл: <b>{fos.SourceFileName}</b></div>");
                        //html.AddFileLink("Дисциплина:", fos.DisciplineName);
                        html.Append($"<div>Дисциплина: <b>{fos.DisciplineName}</b></div>");

                        var fixCompetences1 = Config.FosFixCompetenceTable1;
                        var fixCompetences2 = Config.FosFixCompetenceTable2;
                        var fixPassport = Config.FosFixPassportTable;

                        CurriculumDiscipline discipline = null;
                        var rpd = FindRpd(fos);
                        if (rpd == null) {
                            html.AddError($"Не удалось найти РПД");
                            fixPassport = false;
                            fixCompetences1 = false;
                            fixCompetences2 = false;
                        }
                        else {
                            discipline = FindDiscipline(rpd);
                            if (discipline == null) {
                                fixCompetences1 = false;
                                html.AddError($"Не удалось найти дисциплину [{fos.DisciplineName}] в загруженных учебных планах");
                            }
                        }


                        //var fixEduWorksSummary = Config.RpdFixTableOfEduWorks;
                        //var eduWorks = GetEducationWorks(fos, out _);
                        //if (!(eduWorks?.Any() ?? false)) {
                        //    fixEduWorks = false;
                        //    html.Append($"<div style='color: red'>Не удалось найти учебные работы для дисциплины [{rpd.DisciplineName}] в загруженных учебных планах</div>");
                        //}
                        //if (fos.FormsOfStudy.Count != eduWorks.Count) {
                        //    fixEduWorksSummary = false;
                        //    foreach (var form in fos.FormsOfStudy) {
                        //        if (!eduWorks.ContainsKey(form)) {
                        //            html.Append($"<div style='color: red'>Учебный план для формы обучения [{form.GetDescription()}] не загружен</div>");
                        //        }
                        //    }
                        //}
                        //var fixEduWorkTables = Config.RpdFixFillEduWorkTables;
                        //var eduWorkTableIsFixed = fos.FormsOfStudy.ToDictionary(x => x, x => false);

                        //var eduSummaryTableIsFixed = false; //флаг, что в процессе работы была исправлена сводная таблица учебных работ

                        var setDocProperties = Config.FosFixDocPropertyList?.Where(i => i.IsChecked).ToList();

                        using (var docx = DocX.Load(fos.SourceFileName)) {
                            if (setDocProperties?.Any() ?? false) {
                                foreach (var item in setDocProperties) {
                                    ///docx.CoreProperties[item.Name] = item.Value;
                                    docx.AddCoreProperty(item.Name, item.Value);
                                }
                            }

                            foreach (var table in docx.Tables) {
                                var backup = table.Xml;
                                if (fixCompetences1 && CompetenceMatrix.TestTable(table, out var format) && format == ECompetenceMatrixFormat.Fos21) {
                                    if (RecreateFosTableOfCompetences1(table, rpd, false)) {
                                        html.Append("<div style='color: green'>Таблица компетенций #1 сформирована по матрице компетенций РПД.</div>");
                                    }
                                    else {
                                        html.AddError("Не удалось сформировать обновленную таблицу компетенций #1.");
                                        table.Xml = backup;
                                    }
                                }
                                if (fixCompetences2 && CompetenceMatrix.TestTable(table, out format) && format == ECompetenceMatrixFormat.Fos22) {
                                    if (RecreateFosTableOfCompetences2(table, rpd, false)) {
                                        html.Append("<div style='color: green'>Таблица компетенций #2 сформирована по матрице компетенций РПД.</div>");
                                    }
                                    else {
                                        html.AddError("Не удалось сформировать обновленную таблицу компетенций #2.");
                                        table.Xml = backup;
                                    }
                                }
                                if (fixPassport && TestTableForFosPassport(table, out _, out _)) {
                                    RecreateFosTableOfPassport(table, rpd, true);
                                }
                            }


                            //foreach (var table in docx.Tables) {
                            //    var backup = table.Xml;
                            //    if (fixCompetences && CompetenceMatrix.TestTable(table, out var format) && format == ECompetenceMatrixFormat.Rpd) {
                            //        if (RecreateTableOfCompetences(table, discipline, false)) {
                            //            html.Append("<div style='color: green'>Таблица компетенций сформирована по матрице компетенций.</div>");
                            //        }
                            //        else {
                            //            html.Append("<div style='color: red'>Не удалось сформировать обновленную таблицу компетенций.</div>");
                            //            table.Xml = backup;
                            //        }
                            //    }
                            //    if (fixEduWorksSummary && !eduSummaryTableIsFixed) {
                            //        eduSummaryTableIsFixed |= TestForSummaryTableForEducationalWorks(table, eduWorks, PropertyAccess.Set /*установка значений в таблицу*/);
                            //    }
                            //    if (fixEduWorkTables) {
                            //        //фикс-заполнение таблицы учебных работ для формы обучения
                            //        if (TestForEduWorkTable(table, fos, PropertyAccess.Set, ref evalTools, ref studyResults, out var formOfStudy)) {
                            //            eduWorkTableIsFixed[formOfStudy] = true;
                            //        }
                            //    }
                            //}
                            //if (fixEduWorksSummary) {
                            //    if (eduSummaryTableIsFixed) {
                            //        html.Append("<div style='color: green'>Сводная таблица учебных работ заполнена по учебным планам.</div>");
                            //    }
                            //    else {
                            //        html.Append("<div style='color: red'>Не удалось сформировать сводную таблицу учебных работ.</div>");
                            //        //table.Xml = backup;
                            //    }
                            //}
                            //if (fixEduWorkTables) {
                            //    foreach (var item in eduWorkTableIsFixed) {
                            //        if (item.Value) {
                            //            html.Append($"<div style='color: green'>Таблица учебных работ для формы обучения [{item.Key.GetDescription()}] успешно заполнена.</div>");
                            //        }
                            //        else {
                            //            html.Append($"<div style='color: red'>Таблицу учебных работ для формы обучения [{item.Key.GetDescription()}] не удалось заполнить.</div>");
                            //        }
                            //    }
                            //}

                            //обработка "найти и заменить"
                            var replaceCount = 0;

                            if ((findAndReplaceItems?.Any() ?? false) || App.Config.FosFixResetSelection) {
                                foreach (var par in docx.Paragraphs) {
                                    //поиск и замена
                                    if (findAndReplaceItems?.Any() ?? false) {
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
                                    //очистка цветных выделений
                                    if (App.Config.FosFixResetSelection) {
                                        par.ShadingPattern(new ShadingPattern() { Fill = Color.Transparent, StyleColor = Color.Transparent }, ShadingType.Paragraph);
                                        par.Highlight(Highlight.none);
                                    }
                                }
                                html.Append($"<div>Осуществлено замен в тексте: {replaceCount}</div>");
                            }
                            //extraReplaceItems?.ForEach(x => findAndReplaceItems.Remove(x)); //убираем, т.к. это было нужно только для текущего РПД

                            //фикс разделов по шаблону
                            //if (fixByTemplate) {
                            //    var docxSections = ScanDocxForChapters(docx);
                            //}

                            var fileName = Path.GetFileName(fos.SourceFileName);
                            var newFileName = Path.Combine(targetDir, fileName);
                            if (!Directory.Exists(targetDir)) {
                                Directory.CreateDirectory(targetDir);
                            }
                            docx.SaveAs(newFileName);
                            html.AddFileLink("Итоговый файл:", newFileName);
                            //html.Append($"<div>Итоговый файл: <b>{newFileName}</b></div>");
                        }
                    }
                    else {
                        html.Append($"<div style='color:red'>Файл <b>{fos.SourceFileName}</b> не найден</div>");
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
        /// <param name="useOddHours">можно использовать нечетное кол-во часов</param>
        /// <returns></returns>
        static int[] SplitTime(int total, int count, Random rand,
                               out HashSet<int> luckyNumbers,
                               HashSet<int> includeNumbers = null,
                               HashSet<int> excludeNumbers = null,
                               bool useOddHours = false) {
            var items = new int[count];
            luckyNumbers = includeNumbers?.ToHashSet() ?? new();

            var extraTime = 0;

            double oneItemTime = (double)total / count;
            if (oneItemTime >= 2) { //хватает времени на пару для всех топиков?
                if (oneItemTime % 2 == 0 || useOddHours && oneItemTime % 1 == 0) { //удалось поделить на всех поровну?
                    for (var i = 0; i < count; i++) {
                        items[i] = Convert.ToInt32(oneItemTime);
                    }
                }
                else { //поровну не хватило, надо делить
                    //приводим значение к ближайшему четному (или нечетному, если разрешено) в меньшую сторону
                    var fixedTime = (int)oneItemTime;
                    if (!useOddHours && fixedTime % 2 != 0) fixedTime -= 1;
                    for (int i = 0; i < items.Length; i++) { items[i] = fixedTime; }
                    extraTime = total - fixedTime * count;   //время для распределения
                }
            }
            else {
                extraTime = total;
            }
            if (extraTime > 0) {
                var luckyCount = useOddHours ? extraTime : extraTime >> 1;               //кол-во счастливых топиков, которые получат доп. время
                //var unluckyTopicCount = count - luckyTopicCount;    //оставшиеся - несчастные
                //выбираем случайным образом счастливые топики с учетом уже выбранных номеров, переданных функции
                //var numbers = luckyNumbers?.Take(luckyTopicCount).ToHashSet() ?? new();
                if (excludeNumbers != null && excludeNumbers.Count > count - luckyCount) {
                    excludeNumbers = excludeNumbers.Take(count - luckyCount).ToHashSet();
                }
                while (luckyNumbers.Count < luckyCount) {
                    while (true) {
                        var num = rand.Next(count);
                        if (excludeNumbers?.Contains(num) ?? false) continue; //номер надо исключить
                        if (luckyNumbers.Add(num)) {
                            break;
                        }
                    }
                }
                luckyNumbers = luckyNumbers.Take(luckyCount).ToHashSet();
                foreach (var num in luckyNumbers) {
                    items[num] += useOddHours ? 1 : 2;
                    extraTime -= useOddHours ? 1 : 2;
                }
            }
            if (extraTime > 0 && useOddHours) { //если осталось доп. время и можно применять нечетные часы (так можно для СР)
                //выдадим доп. час случайному топику, кому перепало меньше всего времени
                var minValue = items.Min();
                var indicies = new List<int>();
                for (var i = 0; i < count; i++) {
                    if (items[i] == minValue) {
                        indicies.Add(i);
                    }
                }
                items[indicies[rand.Next(indicies.Count)]] += 1;
            }

            return items;
        }

        /// <summary>
        /// Поделить время обучения
        /// </summary>
        /// <param name="topicCount"></param>
        /// <param name="eduWork"></param>
        /// <returns></returns>
        static (int total, int contact, int lecture, int practical, int selfStudy)[] SplitEduTime(int topicCount, EducationalWork eduWork, Random rand) {
            var topics = new (int total, int contact, int lecture, int practical, int selfStudy)[topicCount];

            //распределяем время на лекции
            var lecItems = SplitTime(eduWork.LectureHours.Value, topicCount, rand, out var luckyNumbers);

            //формируем список счастливых номеров для следующего шага - ими станут текущие несчастливые
            var includeNumbers = new HashSet<int>();
            for (var i = 0; i < topicCount; i++) {
                if (!luckyNumbers.Contains(i)) includeNumbers.Add(i);
            }

            var practicalItems = SplitTime(eduWork.PracticalHours.Value, topicCount, rand, out var luckyNumbers2, includeNumbers: includeNumbers);
            var superLuckyNumbers = luckyNumbers.Intersect(luckyNumbers2).ToHashSet();

            includeNumbers = includeNumbers.Except(luckyNumbers2).ToHashSet();

            var selfStudyItems = SplitTime(eduWork.SelfStudyHours.Value, topicCount, rand, out _, includeNumbers: includeNumbers, excludeNumbers: superLuckyNumbers, true);

            for (int i = 0; i < topicCount; i++) {
                topics[i].lecture = lecItems[i];
                topics[i].practical = practicalItems[i];
                topics[i].selfStudy = selfStudyItems[i];
                topics[i].contact = lecItems[i] + practicalItems[i];
                topics[i].total = lecItems[i] + selfStudyItems[i] + practicalItems[i];
            }

            return topics;
        }

        /// <summary>
        /// Заполнение таблицы учебных работ по заданной форме обучения
        /// </summary>
        /// <param name="workTable"></param>
        /// <param name="rpd"></param>
        /// <param name="formOfStudy"></param>
        static void FillEduWorkTable(DocX docX, Table table, Rpd rpd,
                                     EEduWorkFixType fixTypes,
                                     EFormOfStudy formOfStudy,
                                     ref EEvaluationTool[] evalTools,
                                     ref string[][] studyResults,
                                     decimal maxCompetenceResultsCount,
                                     List<EEvaluationTool> evalTools1stStageItems,
                                     List<EEvaluationTool> evalTools2ndStageItems) {
            var curricula = FindCurricula(rpd);

            if ((curricula?.Any() ?? false) && rpd.EducationalWorks.TryGetValue(formOfStudy, out var eduWork)) {
                var seed = rpd.DisciplineName.ToCharArray().Sum(c => c);
                var rand = new Random(seed);
                var applyFixTime = true; // startCol == 1 && maxColCount == 8;
                //int lastRow = table.RowCount - 2;           //за минусом строки с зачетом/экз. и с итого
                var topicCount = eduWork.TableTopicLastRow - eduWork.TableTopicStartRow + 1;
                Table workTable = table;
                Paragraph parForNewTable = null;

                //требуется полное перестроение
                if (fixTypes.HasFlag(EEduWorkFixType.FullRecreate)) {
                    if (!Templates.Items.TryGetValue(Templates.TEMPLATE_RPD_TABLE_EDU_WORK, out var file)) {
                        throw new Exception($"Невозможно воссоздать таблицу содержания дисциплины, т.к. шаблон [{Templates.TEMPLATE_RPD_TABLE_EDU_WORK}] не найден");
                    }
                    using (var templateDocx = DocX.Load(file)) {
                        if (templateDocx.Tables.Count == 0) {
                            throw new Exception($"Невозможно воссоздать таблицу содержания дисциплины, т.к. в шаблоне [{Templates.TEMPLATE_RPD_TABLE_EDU_WORK}] не найдено таблиц");
                        }
                        var templateTable = templateDocx.Tables.FirstOrDefault();
                        workTable = docX.AddTable(templateTable);
                        //адаптировать свойства EduWork под новую таблицу
                        Templates.AdoptEduWorkForTemplateTable(eduWork, workTable, topicCount);
                    }
                    for (int r = 0; r < topicCount + 2; r++) {
                        workTable.InsertRow();
                    }
                    workTable.Design = TableDesign.TableGrid;

                    parForNewTable = table.Paragraphs.FirstOrDefault().PreviousParagraph;
                    //prevPar.InsertTableAfterSelf(workTable);
                    //workTable.Design = TableDesign.LightShading;

                    //prevPar.InsertParagraphAfterSelf("fuck");
                    //prevPar.InsertParagraphAfterSelf("fuck2");
                    //table.Remove();
                    //var table2 = newTable;
                    //table2.Rows[5].Cells[0].SetText("table<=");
                    //workTable.Rows[6].Cells[0].SetText("newTable<=");
                    //}
                    /*
                    //удаление всех строк ниже заголовка
                    //1. выявим крайний ряд заголовков (в каких-то РПД бывает пустой промежуточный ряд между заголовками и рядами с темами)
                    var headerLastRow = 0;
                    for (var row = 2; row < eduWork.TableTopicStartRow; row++) {
                        if (table.Rows[row].Cells.Any(c => !string.IsNullOrEmpty(c.GetText()))) {
                            headerLastRow = row;
                        }
                        else {
                            break;
                        }
                    }
                    //2. удаление всех рядов ниже заголовка
                    while (table.RowCount != headerLastRow + 1) {
                        table.RemoveRow();
                    }
                    */
                    //добавление новых рядов для тем и двух рядов для контроля и итого

                    //}
                }

                (int total, int contact, int lecture, int practical, int selfStudy)[] topics = null;

                //распределяем время
                if (fixTypes.HasFlag(EEduWorkFixType.Time) && applyFixTime) {
                    topics = SplitEduTime(topicCount, eduWork, rand);

                    //проверка
                    if (topics.Sum(t => t.contact) != (eduWork.ContactWorkHours ?? 0)) {
                        throw new Exception("Сумма времени контактной работы по темам не совпадает с итоговой");
                    }
                    if (topics.Sum(t => t.total) != (eduWork.TotalHours ?? 0) - (eduWork.ControlHours ?? 0)) {
                        throw new Exception("Сумма итогового времени по темам не совпадает с итоговой");
                    }
                    if (topics.Sum(t => t.selfStudy) != (eduWork.SelfStudyHours ?? 0)) {
                        throw new Exception("Сумма времени самостоятельной работы не совпадает с итоговой");
                    }
                }

                //формирование оценочных средств
                if (evalTools == null && fixTypes.HasFlag(EEduWorkFixType.EvalTools)) {
                    evalTools = GetEvalTools(topicCount, rand, evalTools1stStageItems.ToArray(), evalTools2ndStageItems.ToArray());
                }

                var resultCodes = rpd.CompetenceMatrix.GetAllResultCodes().ToList();

                //формирование результатов обучения
                if (studyResults == null && fixTypes.HasFlag(EEduWorkFixType.CompetenceResults)) {
                    studyResults = GetStudyResults(topicCount, resultCodes.ToList(), (int)maxCompetenceResultsCount, rand);
                }

                var colTotal = eduWork.TableStartNumCol;
                var colSubTotal = eduWork.TableHasContactTimeSubtotal ? colTotal + 1 : colTotal;
                var colLecture = colSubTotal + 1;
                var colPractical = colLecture + 1;
                var colSelfStudy = colPractical + 1;

                for (int topic = 0; topic < topicCount; topic++) {
                    var row = topic + eduWork.TableTopicStartRow;
                    //темы
                    if (fixTypes.HasFlag(EEduWorkFixType.FullRecreate)) {
                        workTable.Rows[row].Cells[eduWork.TableColTopic].SetTextFormatted(eduWork.Modules[topic].Topic, fontSize: 11);
                    }
                    //время тем
                    if (fixTypes.HasFlag(EEduWorkFixType.Time) && applyFixTime) {
                        workTable.Rows[row].Cells[colTotal].SetTextFormatted(topics[topic].total, fontSize: 11, alignment: Alignment.center);
                        if (eduWork.TableHasContactTimeSubtotal) {
                            workTable.Rows[row].Cells[colSubTotal].SetTextFormatted(topics[topic].contact, fontSize: 11, alignment: Alignment.center);
                        }
                        workTable.Rows[row].Cells[colLecture].SetTextFormatted(topics[topic].lecture, fontSize: 11, alignment: Alignment.center);
                        workTable.Rows[row].Cells[colPractical].SetTextFormatted(topics[topic].practical, fontSize: 11, alignment: Alignment.center);
                        workTable.Rows[row].Cells[colSelfStudy].SetTextFormatted(topics[topic].selfStudy, fontSize: 11, alignment: Alignment.center);
                    }
                    //оценочные средства
                    if (fixTypes.HasFlag(EEduWorkFixType.EvalTools)) {
                        workTable.Rows[row].Cells[eduWork.TableColEvalTools].SetTextFormatted(evalTools[topic].GetDescription(), fontSize: 11, alignment: Alignment.center);
                        workTable.Rows[row].Cells[eduWork.TableColEvalTools].VerticalAlignment = Xceed.Document.NET.VerticalAlignment.Center;
                    }
                    //результаты компетенций
                    if (fixTypes.HasFlag(EEduWorkFixType.CompetenceResults)) {
                        var codes = studyResults[topic].ToHashSet().ToList();
                        codes.Sort((x1, x2) => resultCodes.IndexOf(x1) - resultCodes.IndexOf(x2));
                        workTable.Rows[row].Cells[eduWork.TableColCompetenceResults].SetTextFormatted(string.Join("\n", codes), fontSize: 11);
                        workTable.Rows[row].Cells[eduWork.TableColCompetenceResults].VerticalAlignment = Xceed.Document.NET.VerticalAlignment.Center;
                    }
                }
                //фикс значения с контролем
                if (fixTypes.HasFlag(EEduWorkFixType.Time) && applyFixTime && eduWork.TableControlRow >= 0) {
                    workTable.Rows[eduWork.TableControlRow].Cells[colTotal].SetTextFormatted(eduWork.ControlHours ?? 0, fontSize: 11);
                }
                //формирование итоговой строки
                if (fixTypes.HasFlag(EEduWorkFixType.Time) && applyFixTime) {
                    var totalRow = eduWork.TableControlRow + 1;
                    workTable.Rows[totalRow].Cells[colTotal].SetTextFormatted(eduWork.TotalHours, fontSize: 11, alignment: Alignment.center);
                    if (eduWork.TableHasContactTimeSubtotal) {
                        workTable.Rows[totalRow].Cells[colSubTotal].SetTextFormatted(eduWork.ContactWorkHours, fontSize: 11, alignment: Alignment.center);
                    }
                    workTable.Rows[totalRow].Cells[colLecture].SetTextFormatted(eduWork.LectureHours, fontSize: 11, alignment: Alignment.center);
                    workTable.Rows[totalRow].Cells[colPractical].SetTextFormatted(eduWork.PracticalHours, fontSize: 11, alignment: Alignment.center);
                    workTable.Rows[totalRow].Cells[colSelfStudy].SetTextFormatted(eduWork.SelfStudyHours, fontSize: 11, alignment: Alignment.center);
                }
                if (fixTypes.HasFlag(EEduWorkFixType.FullRecreate)) {
                    workTable.Rows[eduWork.TableControlRow].Cells[eduWork.TableColTopic].SetTextFormatted(eduWork.ControlForm.GetDescription(), fontSize: 11);
                    workTable.Rows[eduWork.TableControlRow + 1].Cells[eduWork.TableColTopic].SetTextFormatted("Всего по курсу часов:", fontBold: true, fontSize: 11);
                    ResetTableIndentation(workTable);
                }

                //идет формирование новой таблицы из шаблона
                if (parForNewTable != null) {
                    parForNewTable.InsertTableAfterSelf(workTable);
                    table.Remove();
                }
            }
        }

        /// <summary>
        /// Формирование случайных результатов обучения
        /// </summary>
        /// <param name="topicCount"></param>
        /// <param name="rand"></param>
        /// <returns></returns>
        private static string[][] GetStudyResults(int topicCount, List<string> codeList, int maxCount, Random rand) {
            var results = new string[topicCount][];

            var codesToDistrib = codeList.ToList();

            while (codesToDistrib.Any()) {
                var usedNumbers = new HashSet<int>();
                for (var i = 0; i < topicCount; i++) {
                    var max = rand.Next(maxCount) + 1;
                    results[i] = new string[max];
                    for (var r = 0; r < max; r++) {
                        var num = 0;
                        while (true) {
                            if (usedNumbers.Count == codesToDistrib.Count) {
                                usedNumbers.Clear();    //используем повторно полный список результатов компетенций
                            }
                            num = rand.Next(codesToDistrib.Count);
                            if (usedNumbers.Add(num)) break;
                        }
                        results[i][r] = codesToDistrib[num];
                    }
                }
                var usedCodes = new List<string>();
                foreach (var r in results) usedCodes.AddRange(r);
                codesToDistrib = codeList.Except(usedCodes).ToList();
            }

            return results;
        }

        /// <summary>
        /// Формирование случайного списка оценочных средств
        /// </summary>
        /// <param name="topicCount"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private static EEvaluationTool[] GetEvalTools(int topicCount, Random rand, EEvaluationTool[] firstStageTools, EEvaluationTool[] secondStageTools) {
            var tools = new EEvaluationTool[topicCount];
            //var firstTools = new[] { 
            //    EEvaluationTool.Survey, 
            //    EEvaluationTool.Testing 
            //};
            //для первых тем распределим firstTools
            var usedNumbers = new HashSet<int>();
            for (var i = 0; i < firstStageTools.Length; i++) {
                var num = 0;
                while (true) {
                    num = rand.Next(firstStageTools.Length);
                    if (usedNumbers.Add(num)) break;
                }
                tools[i] = firstStageTools[num];
            }
            //var secondTools = new[] {
            //    EEvaluationTool.Essay,
            //    EEvaluationTool.ControlWork,
            //    EEvaluationTool.Paper,
            //    EEvaluationTool.Presentation
            //};
            //вторая порция средств
            usedNumbers.Clear();
            for (var i = 0; i < Math.Min(topicCount - firstStageTools.Length, secondStageTools.Length); i++) {
                var num = 0;
                while (true) {
                    num = rand.Next(secondStageTools.Length);
                    if (usedNumbers.Add(num)) break;
                }
                tools[i + firstStageTools.Length] = secondStageTools[num];
            }
            //третья и последующая порция
            var thirdTools = firstStageTools.Concat(secondStageTools).ToArray(); //Enum.GetValues(typeof(EEvaluationTool)).Cast<EEvaluationTool>().ToArray();
            var topicIdx = firstStageTools.Length + secondStageTools.Length;
            usedNumbers.Clear();

            while (topicIdx < topicCount) {
                if (usedNumbers.Count == thirdTools.Length) {
                    usedNumbers.Clear(); //следующий цикл
                }
                var num = 0;
                while (true) {
                    num = rand.Next(thirdTools.Length);
                    if (usedNumbers.Add(num)) break;
                }
                tools[topicIdx++] = thirdTools[num];
            }

            return tools;
        }

        /// <summary>
        /// Проверка на таблицу учебных работ по форме обучения
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rpd"></param>
        /// <returns></returns>
        internal static bool TestForEduWorkTable(DocX docX, Table table, Rpd rpd, PropertyAccess propAccess,
                                                 EEduWorkFixType fixTypes,
                                                 ref EEvaluationTool[] evalTools,
                                                 ref string[][] studyResults,
                                                 decimal maxCompetenceResultsCount,
                                                 List<EEvaluationTool> evalTools1stStageItems,
                                                 List<EEvaluationTool> evalTools2ndStageItems,
                                                 out EFormOfStudy formOfStudy) {
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
                        rpd.EducationalWorks.TryGetValue(form, out var eduWork) &&
                        regex.IsMatch(par.Text)) {

                        //определяем начальный значащие ряд и колонку
                        var startRow = 3;
                        var startNumCol = -1;
                        var firstNumCol = -1;
                        var lastNumCol = -1;
                        //var numColCount = 0;    //кол-во числовых ячеек
                        while (startRow < table.RowCount) {
                            //numColCount = 0;
                            firstNumCol = -1;
                            lastNumCol = -1;
                            var hasNumCells = false;
                            for (var col = 0; col < table.Rows[startRow].Cells.Count; col++) {
                                var cell = table.Rows[startRow].Cells[col];
                                //if (cell.GridSpan > 0) { //первый ряд, который нам нужен - ряд без объединений ячеек
                                //    startRow++;
                                //    continue;
                                //}
                                var text = cell.GetText();
                                if (int.TryParse(text, out _)) {
                                    hasNumCells = true;
                                    if (firstNumCol < 0) firstNumCol = col;
                                    if (col > lastNumCol) lastNumCol = col;
                                    if (col > 0) {
                                        if (startNumCol < 0) startNumCol = col;
                                        //numColCount++;
                                    }
                                }
                            }
                            if (hasNumCells) {
                                break;
                            }
                            startRow++;
                        }
                        eduWork.TableTopicStartRow = startRow;
                        eduWork.TableStartNumCol = startNumCol;
                        eduWork.TableColTopic = startNumCol - 1;
                        eduWork.TableMaxColCount = lastNumCol + 3;
                        eduWork.TableColEvalTools = lastNumCol + 1;
                        eduWork.TableColCompetenceResults = lastNumCol + 2;

                        //тестируем на кол-во ячеек с числовыми значениями: должно быть 4 или 5 (когда есть подитог по КР)
                        eduWork.TableHasContactTimeSubtotal = lastNumCol - firstNumCol + 1 == 5;

                        //попытка определить последний ряд топика (в таблице могут быть строки с "зачет/экзамен", а также "курсовая", "всего за X семестр")
                        //var nonTopicRowIdx = table.RowCount - 3;
                        for (var row = table.RowCount - 2; row >= startRow; row--) {
                            var topic = table.Rows[row].Cells[eduWork.TableColTopic].GetText();
                            if (!string.IsNullOrEmpty(topic)) {
                                if (Regex.IsMatch(topic.ToLower(), @"экзамен|зач[е,ё]т")) {
                                    eduWork.TableControlRow = row;
                                }
                                if (!Regex.IsMatch(topic.ToLower(), @"экзамен|зач[е,ё]т|курсовая|всего за")) {
                                    eduWork.TableTopicLastRow = row;
                                    break;
                                }
                            }
                        }

                        if (propAccess == PropertyAccess.Get) {
                            rpd.EducationalWorks[form].Table ??= table;
                        }
                        else {
                            //простановка времени по темам
                            FillEduWorkTable(docX, table, rpd, fixTypes, form, ref evalTools, ref studyResults,
                                             maxCompetenceResultsCount,
                                             evalTools1stStageItems,
                                             evalTools2ndStageItems);
                        }
                        formOfStudy = form;
                        result = true;
                        break;
                    }
                }
                if (result) break;
            }

            return result;
        }

        /// <summary>
        /// Вернуть список возможных предшествующих дисциплин (список далее будет фильтроваться YaGPT)
        /// </summary>
        /// <param name="discipline"></param>
        /// <returns></returns>
        internal static List<CurriculumDiscipline> GetPossiblePrevDisciplines(Rpd rpd) {
            var discList = new List<CurriculumDiscipline>();

            var curricula = App.FindCurricula(rpd);
            if (curricula != null && curricula.Any()) {
                var curr = curricula.First().Value;

                var disc = curr.FindDiscipline(rpd.DisciplineName);
                if (disc != null && disc.StartSemesterIdx >= 0) {
                    //ищем все дисциплины, которые стартуют раньше нужной или одновременно
                    discList = curr.Disciplines.Values.Where(d => d != disc &&
                                                                  d.StartSemesterIdx <= disc.StartSemesterIdx &&
                                                                  d.Type != EDisciplineType.Optional).ToList();
                }
            }

            return discList;
        }

        /// <summary>
        /// Нормализация имени ("Основы микро- и макроэкономики" -> "ОСНОВЫМИКРОИМАКРОЭКОНОМИКИ")
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        internal static string NormalizeName(string name) {
            if (!string.IsNullOrEmpty(name)) {
                var newName = string.Join("", name.Split(' ', '-', '(', ')', '/', '.', ';', ':', '[', ']'));
                return newName.ToUpper().Replace('Ё', 'Е');
            }
            else {
                return name;
            }
        }

        /// <summary>
        /// Вернуть список возможных последующих дисциплин (список далее будет фильтроваться YaGPT)
        /// </summary>
        /// <param name="discipline"></param>
        /// <returns></returns>
        internal static List<CurriculumDiscipline> GetPossibleNextDisciplines(Rpd rpd) {
            var discList = new List<CurriculumDiscipline>();

            var curricula = App.FindCurricula(rpd);
            if (curricula != null && curricula.Any()) {
                var curr = curricula.First().Value;

                var disc = curr.FindDiscipline(rpd.DisciplineName);
                if (disc != null && disc.LastSemesterIdx >= 0) {
                    //ищем все дисциплины, которые стартуют раньше нужной или одновременно
                    discList = curr.Disciplines.Values.Where(d => d != disc &&
                                                                  d.LastSemesterIdx >= disc.LastSemesterIdx &&
                                                                  d.Type != EDisciplineType.Optional).ToList();
                }
            }

            return discList;
        }

        /// <summary>
        /// Сохранить стор
        /// </summary>
        public static void SaveStore(EStoreElements elements) {
            if (elements.HasFlag(EStoreElements.Rpd)) {
                var storeFileRpd = Path.Combine(Environment.CurrentDirectory, Store.FILE_NAME_RPD);
                if (m_config != null) {
                    var json = JsonSerializer.Serialize(m_store.RpdDic, m_jsonOptions);
                    File.WriteAllText(storeFileRpd, json, Encoding.UTF8);
                }
            }
            if (elements.HasFlag(EStoreElements.Fos)) {
                var storeFileFos = Path.Combine(Environment.CurrentDirectory, Store.FILE_NAME_FOS);
                if (m_config != null) {
                    var json = JsonSerializer.Serialize(m_store.FosDic, m_jsonOptions);
                    File.WriteAllText(storeFileFos, json, Encoding.UTF8);
                }
            }
            if (elements.HasFlag(EStoreElements.Curricula)) {
                var storeFileCurricula = Path.Combine(Environment.CurrentDirectory, Store.FILE_NAME_CURRICULA);
                if (m_config != null) {
                    var json = JsonSerializer.Serialize(m_store.CurriculaDic, m_jsonOptions);
                    File.WriteAllText(storeFileCurricula, json, Encoding.UTF8);
                }
            }
        }

        /// <summary>
        /// Загрузить стор
        /// </summary>
        public static void LoadStore(EStoreElements elements = EStoreElements.All) {
            m_store ??= new();

            //РПД
            if (elements.HasFlag(EStoreElements.Rpd)) {
                var storeFileRpd = Path.Combine(Environment.CurrentDirectory, Store.FILE_NAME_RPD);
                if (File.Exists(storeFileRpd)) {
                    var json = File.ReadAllText(storeFileRpd);
                    m_store.RpdDic = JsonSerializer.Deserialize<Dictionary<string, Rpd>>(json, m_jsonOptions);
                }
                else {
                    m_store.RpdDic = [];
                }
            }
            //ФОС
            if (elements.HasFlag(EStoreElements.Fos)) {
                var storeFileFos = Path.Combine(Environment.CurrentDirectory, Store.FILE_NAME_FOS);
                if (File.Exists(storeFileFos)) {
                    var json = File.ReadAllText(storeFileFos);
                    m_store.FosDic = JsonSerializer.Deserialize<Dictionary<string, Fos>>(json, m_jsonOptions);
                }
                else {
                    m_store.FosDic = [];
                }
            }
            //УП
            if (elements.HasFlag(EStoreElements.Curricula)) {
                var storeFileCurricula = Path.Combine(Environment.CurrentDirectory, Store.FILE_NAME_CURRICULA);
                if (File.Exists(storeFileCurricula)) {
                    var json = File.ReadAllText(storeFileCurricula);
                    m_store.CurriculaDic = JsonSerializer.Deserialize<Dictionary<string, Curriculum>>(json, m_jsonOptions);
                    foreach (var item in m_store.CurriculaDic.Values) {
                        if (item.Disciplines != null) {
                            foreach (var disc in item.Disciplines.Values) {
                                disc.Curriculum = item; //восст. ссылки на родительский УП
                            }
                        }
                    }
                }
                else {
                    m_store.CurriculaDic = [];
                }
            }
        }

        /// <summary>
        /// Добавить/обновить загруженные РПД в стор
        /// </summary>
        //public static void AddLoadedRpdToStore() {
        //    foreach (var rpd in m_rpdDic) {
        //        m_store.RpdDic[rpd.Key] = rpd.Value;
        //    }
        //    SaveStore();
        //}

        ///// <summary>
        ///// Добавить/обновить загруженные ФОС в стор
        ///// </summary>
        //public static void AddLoadedFosToStore() {
        //    foreach (var fos in m_fosDic) {
        //        m_store.FosDic[fos.Key] = fos.Value;
        //    }
        //    SaveStore();
        //}

        /// <summary>
        /// Вернуть случайное значение элементов из списка (от, до)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="fromCount"></param>
        /// <param name="toCount"></param>
        /// <returns></returns>
        public static List<T> TakeRandom<T>(this List<T> list, int fromCount, int toCount) {
            var count = fromCount + new Random().Next(toCount - fromCount + 1);
            return list.Take(count).ToList();
        }

        public static Task ForEachAsync<T>(this IEnumerable<T> sequence, Func<T, Task> action) {
            return Task.WhenAll(sequence.Select(action));
        }


        /// <summary>
        /// Проверка таблицы на матрицу компетенций
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rpd"></param>
        public static bool TestForTableOfCompetenceMatrix(Table table,
                                                          ECompetenceMatrixFormat format,
                                                          out CompetenceMatrix matrix,
                                                          out List<Error> errors) {
            matrix = new CompetenceMatrix() {
                Items = [],
                Errors = new(),
            };
            errors = null;

            var result = CompetenceMatrix.TryParseTable(table, format, matrix);
            errors = matrix.Errors?.GetCopy();

            if (!result) {
                matrix = null;
            }

            return result;
        }

        //функция для нормализации текста: убираем лишние пробелы, приводим к UpperCase, удаляем цифры, тире, точки
        static string NormalizeText(string text) => string.Concat(string.Join(" ", text.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                                                                  .Where(c => c != '-' && c != '.' && (c < '0' || c > '9'))).ToUpper().Trim();

        /// <summary>
        /// Проверка таблицы на паспорт ФОСа (3. Паспорт фонда оценочных средств текущего контроля, соотнесённых с индикаторами достижения компетенций)
        /// </summary>
        /// <param name="table"></param>
        /// <param name="passport"></param>
        /// <returns></returns>
        public static bool TestTableForFosPassport(Table table, out List<StudyModule> passport, out List<string> errors) {
            errors = [];
            passport = null;

            try {
                //ожидаем колонки
                //№ п/п	- Контролируемые модули, разделы (темы) дисциплины - Код контролируемого индикатора достижения компетенции - Наименование оценочного средства
                if (table.RowCount > 2 && table.ColumnCount >= 4) {
                    var headerRow = table.Rows[0];
                    var header1 = headerRow.Cells[1].GetText(); //Контролируемые модули, разделы (темы) дисциплины?
                    var header2 = headerRow.Cells[2].GetText(); //Код контролируемого индикатора достижения компетенции?
                    var header3 = headerRow.Cells[3].GetText(); //Наименование оценочного средства?
                    //проверка заголовков
                    if (m_regexFosPassportHeaderTopic.IsMatch(header1) &&
                        m_regexFosPassportHeaderIndicator.IsMatch(header2) &&
                        m_regexFosPassportHeaderEvalTool.IsMatch(header3)) {
                        passport = [];

                        //нормализуем список оценочных средств
                        var evalToolDic = Enum.GetValues(typeof(EEvaluationTool)).Cast<EEvaluationTool>().ToDictionary(x => x.GetDescription().ToUpper(), x => x);

                        for (var row = 1; row < table.RowCount; row++) {
                            //var evalTools = table.Rows[row].Cells[3].GetText(",").Split(',', '\n', ';');
                            HashSet<EEvaluationTool> tools = [];
                            foreach (var t in table.Rows[row].Cells[3].GetText(",").Split(',', '\n', ';')) {
                                //убираем лишние пробелы
                                if (evalToolDic.TryGetValue(NormalizeText(t), out var tool)) {
                                    //var normalizedText = string.Join(" ", t.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)).ToUpper();
                                    //if (Enum.TryParse(normalizedText, true, out EEvaluationTool tool)) {
                                    tools.Add(tool);
                                }
                                else {
                                    errors.Add($"Не удалось определить тип оценочного средства - {t}");
                                }
                            }

                            passport.Add(new() {
                                Topic = table.Rows[row].Cells[1].GetText(),
                                CompetenceResultCodes = [.. table.Rows[row].Cells[2].GetText(",").Split(',', '\n', ';')],
                                EvaluationTools = tools
                            });
                        }
                    }
                }
            }
            catch (Exception ex) {
                errors.Add($"{ex.Message}\r\n{ex.StackTrace}");
            }

            return passport != null;
        }

        /// <summary>
        /// Удалить РПД из загруженных
        /// </summary>
        /// <param name="rpd"></param>
        internal static void RemoveRpd(Rpd rpd) {
            if (rpd != null) {
                m_rpdDic.TryRemove(rpd.SourceFileName, out _);
            }
        }

        /// <summary>
        /// Удалить список РПД из загруженных
        /// </summary>
        /// <param name="rpdList"></param>
        internal static void RemoveRpd(List<Rpd> rpdList) {
            if (rpdList?.Any() ?? false) {
                foreach (var rpd in rpdList) {
                    m_rpdDic.TryRemove(rpd.SourceFileName, out _);
                }
            }
        }

        /// <summary>
        /// Удалить из списка ряд УП
        /// </summary>
        /// <param name="curriculaList"></param>
        internal static void RemoveCurricula(List<Curriculum> curriculaList) {
            if (curriculaList?.Any() ?? false) {
                foreach (var curriculum in curriculaList) {
                    m_curriculumDic.TryRemove(curriculum.SourceFileName, out _);
                }
            }
        }

        /// <summary>
        /// Удалить из списка УП
        /// </summary>
        /// <param name="curriculaList"></param>
        internal static void RemoveCurriculum(Curriculum curriculum) {
            if (curriculum != null) {
                m_curriculumDic.TryRemove(curriculum.SourceFileName, out _);
            }
        }
    }
}
