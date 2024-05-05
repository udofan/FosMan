using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.Eventing.Reader;
using System.DirectoryServices.ActiveDirectory;
using System.Drawing;
using System.Drawing.Design;
using System.Linq;
using System.Reflection;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace FosMan {
    /// <summary>
    /// Загруженные данные
    /// </summary>
    static internal class App {
        static CompetenceMatrix m_competenceMatrix = null;
        static Dictionary<string, Curriculum> m_curriculumDic = [];
        static Dictionary<string, Rpd> m_rpdDic = [];
        static Dictionary<string, CurriculumGroup> m_curriculumGroupDic = [];

        public static CompetenceMatrix CompetenceMatrix { get => m_competenceMatrix; }

        public static Dictionary<string, Rpd> RpdList { get => m_rpdDic; }

        /// <summary>
        /// Учебные планы в сторе
        /// </summary>
        public static Dictionary<string, Curriculum> Curricula { get => m_curriculumDic; }
        /// <summary>
        /// Группы УП
        /// </summary>
        public static Dictionary<string, CurriculumGroup> CurriculumGroups { get => m_curriculumGroupDic; }

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
                        Profile = curriculum.Profile
                    };
                    m_curriculumGroupDic[curriculum.GroupKey] = group;
                }
                group.AddCurriculum(curriculum);
                result = true;
            }
            return result;
        }

        static public bool AddRpd(Rpd rpd) {
            return m_rpdDic.TryAdd(rpd.SourceFileName, rpd);
        }

        static public void SetCompetenceMatrix(CompetenceMatrix matrix) {
            m_competenceMatrix = matrix;
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

        static void AddDiv(this StringBuilder report, string content) {
            report.Append($"<div>{content}</div>");
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
                                         ref int pos, ref int errorCount, string description, Func< EducationalWork, (bool result, string msg)> func) {
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
        public static void CheckRdp(out string htmlReport) {
            htmlReport = string.Empty;

            StringBuilder html = new("<html><body><h2>Отчёт по проверке РПД</h2>");
            StringBuilder toc = new("<div><ul>");
            StringBuilder rep = new("<div>");
            var idx = 0;
            foreach (var rpd in m_rpdDic.Values) {
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
                var curriculums = FindCurriculums(rpd.DirectionCode, rpd.Profile, rpd.Department);
                if (curriculums != null && curriculums.Any()) {
                    //проверка на отсутствие УПов
                    var missedFormsOfStudy = rpd.FormsOfStudy.ToHashSet();
                    missedFormsOfStudy.RemoveWhere(x => curriculums.ContainsKey(x));
                    foreach (var item in missedFormsOfStudy) {
                        rep.Append("<p />");
                        rep.AddDiv($"Форма обучения: <b>{item.GetDescription()}</b>");
                        rep.AddDiv($"Файл УП: <b><span style='color: red'>Не найден УП для данной формы обучения</span></b>");
                        rep.AddError("Проверка РПД по УП невозможна.");
                        //rep.AddError($"Не найден УП для формы обучения <b>{item.GetDescription()}</b>.");
                        errorCount++;
                    }

                    //поиск дисциплины в УПах
                    foreach (var curriculum in curriculums) {
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
                                    var result = discipline.TotalByPlanHours == eduWork.TotalHours;
                                    var msg = result ? "" : $"Итоговое время [{discipline.TotalByPlanHours}] не соответствует УП (д.б. {eduWork.TotalHours}).";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Время контактной работы", (eduWork) => {
                                    var result = discipline.TotalContactWorkHours == eduWork.ContactWorkHours;
                                    var msg = result ? "" : $"Время контактной работы [{discipline.TotalContactWorkHours}] не соответствует УП (д.б. {eduWork.ContactWorkHours}).";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Время контроля", (eduWork) => {
                                    var result = discipline.TotalControlHours == eduWork.ControlHours;
                                    var msg = result ? "" : $"Время контроля [{discipline.TotalControlHours}] не соответствует УП (д.б. {eduWork.ControlHours}).";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Время самостоятельной работы", (eduWork) => {
                                    var result = discipline.TotalSelfStudyHours == eduWork.SelfStudyHours;
                                    var msg = result ? "" : $"Время самостоятельных работ [{discipline.TotalSelfStudyHours}] не соответствует УП (д.б. {eduWork.SelfStudyHours}).";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Время практических работ", (eduWork) => {
                                    var practicalHours = discipline.Semesters.Sum(s => s.PracticalHours);
                                    var result = practicalHours == eduWork.PracticalHours;
                                    var msg = result ? "" : $"Время практических работ [{practicalHours}] не соответствует УП (д.б. {eduWork.PracticalHours}).";
                                    return (result, msg);
                                });
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Время лекций", (eduWork) => {
                                    var lectureHours = discipline.Semesters.Sum(s => s.LectureHours);
                                    var result = lectureHours == eduWork.LectureHours;
                                    var msg = result ? "" : $"Время лекций [{lectureHours}] не соответствует УП (д.б. {eduWork.LectureHours}).";
                                    return (result, msg);
                                });
                                //проверка итогового контроля
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Тип итогового контроля", (eduWork) => {
                                    var error = eduWork.ControlForm == EControlForm.Exam && (!discipline.ControlFormExamHours.HasValue || discipline.ControlFormExamHours == 0) ||
                                                eduWork.ControlForm == EControlForm.Test && (!discipline.ControlFormTestHours.HasValue || discipline.ControlFormTestHours == 0) ||
                                                eduWork.ControlForm == EControlForm.TestWithAGrade && (!discipline.ControlFormTestWithAGradeHours.HasValue || discipline.ControlFormTestWithAGradeHours == 0);
                                    var msg = error ? $"Для типа контроля [{eduWork.ControlForm.GetDescription()}] в УП не определено значение." : "";
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

                            if (checkCompetences) {
                                ApplyDisciplineCheck(curriculum.Key, rpd, discipline, table, ref checkPos, ref errorCount, "Проверка матрицы компетенций", (eduWork) => {
                                    var achiCodeList = new List<string>();
                                    rpd.CompetenceMatrix.Items.ForEach(x => achiCodeList.AddRange(x.Achievements.Select(a => a.Code)));
                                    var content = "";
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

                                        if (content.Length > 0) content += "; ";
                                        content += elem;
                                    }
                                    foreach (var missedCode in achiCodeList.Except(discipline.CompetenceList)) {
                                        matrixError = true;
                                        if (content.Length > 0) content += "; ";
                                        content += $"<span style='color: red; font-decoration: italic'>{missedCode}??</span>"; ;
                                    }
                                    var msg = matrixError ? $"Выявлено несоответствие компетенций: {content}" : "";
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
        private static Dictionary<EFormOfStudy, Curriculum> FindCurriculums(string directionCode, string profile, string department) {
            if (!string.IsNullOrEmpty(directionCode) && !string.IsNullOrEmpty(profile) && !string.IsNullOrEmpty(department)) {
                var items = m_curriculumDic.Values.Where(c => string.Compare(c.DirectionCode, directionCode, true) == 0 &&
                                                              string.Compare(c.Profile, profile, true) == 0 &&
                                                              string.Compare(c.Department, department, true) == 0);
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
                                                    Action<int, CurriculumDiscipline> progressAction, out List<string> errors) {
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
                        i++;
                        progressAction?.Invoke(i, disc);
                        string fileName;
                        if (!string.IsNullOrEmpty(fileNameTemplate)) {
                            fileName = Regex.Replace(fileNameTemplate, @"({[^}]+})", (m) => (disc.GetProperty(m.Value.Trim('{', '}'))?.ToString() ?? m.Value));
                        }
                        else {  //если шаблон имени файла не задан
                            fileName = $"РПД_{disc.Index}_{disc.Name}.docx";
                        }
                        //создание файла
                        var targetFile = Path.Combine(targetDir, fileName);
                        File.Copy(rpdTemplate, targetFile, true);
                        using (var docx = DocX.Load(targetFile)) {
                            foreach (var par in docx.Paragraphs.ToList()) {
                                var replaceOptions = new FunctionReplaceTextOptions() {
                                    ContainerLocation = ReplaceTextContainer.All,
                                    FindPattern = "{[^}]+}",
                                    RegexMatchHandler = m => {
                                        var replaceValue = m;
                                        var propName = m.Trim('{', '}');
                                        var groupProp = curriculumGroup.GetProperty(propName);
                                        if (groupProp != null) {
                                            replaceValue = groupProp.ToString();
                                        }
                                        else {
                                            var discProp = disc.GetProperty(propName);
                                            if (discProp != null) {
                                                replaceValue = discProp.ToString();
                                            }
                                            else {
                                                if (TryInsertSpecialObjectInRpd(propName, par)) {
                                                    replaceValue = "";
                                                }
                                            }
                                        }
                                        return replaceValue;
                                    }
                                };
                                par.ReplaceText(replaceOptions);

                                //par.ReplaceText(@"{[^}]+}", m => {
                                //});                                //par.ReplaceText(@"{[^}]+}", m => {
                                //});
                            }
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
        /// Попытка вставить специальный объект в параграф
        /// </summary>
        /// <param name="propName"></param>
        /// <param name="par"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        private static bool TryInsertSpecialObjectInRpd(string propName, Paragraph par) {
            var result = false;
            switch (propName) {
                case "TableCompetences":
                    var newTable = par.InsertTableAfterSelf(5, 3);
                    //newTable.ad
                    result = true;
                    break;
                case "TableEducationWorks":
                    var newTable2 = par.InsertTableAfterSelf(5, 4);
                    result = true;
                    break;
                default:
                    break;
            }

            return result;
        }
    }
}
