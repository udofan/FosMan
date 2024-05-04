using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.DirectoryServices.ActiveDirectory;
using System.Drawing;
using System.Drawing.Design;
using System.Linq;
using System.Reflection;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    /// <summary>
    /// Загруженные данные
    /// </summary>
    static internal class App {
        static CompetenceMatrix m_competenceMatrix = null;
        static Dictionary<string, Curriculum> m_curriculumDic = [];
        static Dictionary<string, Rpd> m_rpdDic = [];

        public static CompetenceMatrix CompetenceMatrix { get => m_competenceMatrix; }

        public static Dictionary<string, Rpd> RpdList { get => m_rpdDic; }

        public static Dictionary<string, Curriculum> Curricula { get => m_curriculumDic; }

        public static bool HasCurriculumFile(string fileName) => m_curriculumDic.ContainsKey(fileName);

        public static bool HasRpdFile(string fileName) => m_rpdDic.ContainsKey(fileName);

        static public bool AddCurriculum(Curriculum curriculum) {
            return m_curriculumDic.TryAdd(curriculum.SourceFileName, curriculum);
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

        static void AddTocElement(this StringBuilder toc, string element, string anchor, bool errorMark) {
            var color = errorMark ? "red" : "green";
            toc.Append($"<li><a href='#{anchor}'><span style='color:{color};'>{element}</span></a></li>");
        }

        /// <summary>
        /// Проверить загруженные РПД
        /// </summary>
        public static void CheckRdp(out string htmlReport) {
            htmlReport = string.Empty;

            StringBuilder html = new("<html><body>");
            StringBuilder toc = new("<div><ul>");
            StringBuilder rep = new("<div>");
            var idx = 0;
            foreach (var rpd in m_rpdDic.Values) {
                var anchor = $"rpd{idx}";
                idx++;
                rpd.ExtraErrors = [];
                var hasErrors = false;

                rep.Append($"<div id='{anchor}'><h3>{rpd.DisciplineName}</h3>");
                rep.AddDiv(rpd.SourceFileName);
                //ищем Учебные планы
                var curriculums = FindCurriculums(rpd.DirectionCode, rpd.Profile, rpd.Department);
                if (curriculums != null && curriculums.Any()) {
                    //проверка на отсутствие УПов
                    var missedFormsOfStudy = rpd.FormsOfStudy.ToHashSet();
                    missedFormsOfStudy.RemoveWhere(x => curriculums.ContainsKey(x));
                    foreach (var item in missedFormsOfStudy) {
                        rep.AddError($"Не найден УП для формы обучения [{item.GetDescription()}].");
                        hasErrors = true;
                    }

                    //поиск дисциплины в УПах
                    foreach (var curriculum in curriculums) {
                        var discipline = curriculum.Value.FindDiscipline(rpd.DisciplineName);
                        if (discipline != null) {
                            //проверка времен учебной работы
                            var eduWork = rpd.EducationalWorks[curriculum.Key];
                            if (discipline.TotalByPlanHours != eduWork.TotalHours) {
                                hasErrors = true;
                                rep.AddError($"Форма обучения [{curriculum.Key.GetDescription()}: итоговое время [{discipline.TotalByPlanHours}] не соответствует УП (д.б. {eduWork.TotalHours}).");
                            }
                            if (discipline.TotalContactWorkHours != eduWork.ContactWorkHours) {
                                hasErrors = true;
                                rep.AddError($"Форма обучения [{curriculum.Key.GetDescription()}: время контактной работы [{discipline.TotalContactWorkHours}] не соответствует УП (д.б. {eduWork.ContactWorkHours}).");
                            }
                            if (discipline.TotalControlHours != eduWork.ControlHours) {
                                hasErrors = true;
                                rep.AddError($"Форма обучения [{curriculum.Key.GetDescription()}: время контроля [{discipline.TotalControlHours}] не соответствует УП (д.б. {eduWork.ControlHours}).");
                            }
                            if (discipline.TotalSelfStudyHours!= eduWork.SelfStudyHours) {
                                hasErrors = true;
                                rep.AddError($"Форма обучения [{curriculum.Key.GetDescription()}: время самостоятельных работ [{discipline.TotalSelfStudyHours}] не соответствует УП (д.б. {eduWork.SelfStudyHours}).");
                            }
                            var practicalHours = discipline.Semesters.Sum(s => s.PracticalHours);
                            if (practicalHours != eduWork.PracticalHours) {
                                hasErrors = true;
                                rep.AddError($"Форма обучения [{curriculum.Key.GetDescription()}: время практических работ [{practicalHours}] не соответствует УП (д.б. {eduWork.PracticalHours}).");
                            }
                            var lectureHours = discipline.Semesters.Sum(s => s.LectureHours);
                            if (lectureHours != eduWork.LectureHours) {
                                hasErrors = true;
                                rep.AddError($"Форма обучения [{curriculum.Key.GetDescription()}: время лекций [{lectureHours}] не соответствует УП (д.б. {eduWork.LectureHours}).");
                            }
                            //проверка итогового контроля
                            if (eduWork.ControlForm == EControlForm.Exam && 
                                (!discipline.ControlFormExamHours.HasValue || discipline.ControlFormExamHours == 0)) {
                                hasErrors = true;
                                rep.AddError($"Форма обучения [{curriculum.Key.GetDescription()}: выявлено несоответствие типа итогового контроля - " +
                                             $"ожидается [{eduWork.ControlForm.GetDescription()}], а в УП для экзамена время не определено.");
                            }
                            if (eduWork.ControlForm == EControlForm.Test &&
                                (!discipline.ControlFormTestHours.HasValue || discipline.ControlFormTestHours == 0)) {
                                hasErrors = true;
                                rep.AddError($"Форма обучения [{curriculum.Key.GetDescription()}: выявлено несоответствие типа итогового контроля - " +
                                             $"ожидается [{eduWork.ControlForm.GetDescription()}], а в УП для зачета время не определено.");
                            }
                            if (eduWork.ControlForm == EControlForm.TestWithAGrade &&
                                (!discipline.ControlFormTestWithAGradeHours.HasValue || discipline.ControlFormTestWithAGradeHours == 0)) {
                                hasErrors = true;
                                rep.AddError($"Форма обучения [{curriculum.Key.GetDescription()}: выявлено несоответствие типа итогового контроля - " +
                                             $"ожидается [{eduWork.ControlForm.GetDescription()}], а в УП для зачета с оценкой время не определено.");
                            }
                            //проверка компетенций
                            var checkCompetences = true;
                            if (!(discipline.CompetenceList?.Any() ?? false)) {
                                hasErrors = true;
                                rep.AddError($"Форма обучения [{curriculum.Key.GetDescription()}: не удалось определить матрицу компетенций в УП.");
                                checkCompetences = false;
                            }
                            if (!(rpd.CompetenceMatrix?.IsLoaded ?? false)) {
                                hasErrors = true;
                                rep.AddError("В РПД не обнаружена матрица компетенций - проверка компетенций невозможна.");
                                checkCompetences = false;
                            }
                            if (!(m_competenceMatrix?.IsLoaded ?? false)) {
                                hasErrors = true;
                                rep.AddError("Матрица компетенций не загружена в программу - проверка компетенций невозможна.");
                                checkCompetences = false;
                            }
                            if (checkCompetences) {
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
                                if (matrixError) {
                                    hasErrors = true;
                                    rep.AddError($"Форма обучения [{curriculum.Key.GetDescription()}: выявлено несоответствие компетенций: {content}");
                                }
                            }
                        }
                        else {
                            hasErrors = true;
                            rep.AddError($"Не удалось найти дисциплину в учебном плане [{curriculum.Value.SourceFileName}].");
                        }
                    }
                }
                else {
                    rep.AddError("Не удалось найти учебные планы. Добавление УП осуществляется на вкладке \"Учебные планы\".");
                    hasErrors = true;
                }
                if (!hasErrors) {
                    rep.AddDiv("Ошибок не обнаружено.");
                }
                rep.Append("</div>");

                toc.AddTocElement(rpd.DisciplineName, anchor, hasErrors);
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
    }
}
