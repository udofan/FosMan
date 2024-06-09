using FastMember;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Numerics;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    /// <summary>
    /// Форма контроля
    /// </summary>
    public enum EControlForm {
        [Description("Экзамен")]
        Exam,           //экзамен
        [Description("Зачет")]
        Test,           //зачет
        [Description("Зачет с оценкой")]
        TestWithAGrade, //зачет с оценкой
        [Description("Контрольная работа")]
        CourseWork,     //курсовая работа
        [Description("НЕИЗВЕСТНО")]
        Unknown         //определить НЕ УДАЛОСЬ
    }

    /// <summary>
    /// Тип дисциплины
    /// </summary>
    public enum EDisciplineType {
        [Description("Обязательная")]
        Required,       //обязательная
        [Description("Вариативная")]
        Variable,
        [Description("По выбору")]
        ByChoice,       //по выбору
        [Description("Факультативная")]
        Optional,       //факультативная
        [Description("НЕИЗВЕСТНО")]
        Unknown         //опередить НЕ УДАЛОСЬ
    }

    /// <summary>
    /// Описание дисциплины
    /// </summary>
    internal class CurriculumDiscipline : BaseObj {
        public const int SEMESTER_COUNT = 10;

        static Regex m_regexTestTypeRequired = new(@"^[^\.]+\.О\.\d+", RegexOptions.Compiled);      //обязательная
        static Regex m_regexTestTypeVariable = new(@"^[^\.]+\.В\.\d+", RegexOptions.Compiled);      //обязательная часть, формируемая участниками
        static Regex m_regexTestTypeByChoice = new(@"^[^\.]+\.В\.ДВ\.\d+", RegexOptions.Compiled);  //выбору части, формируемой...
        static Regex m_regexTestTypeOptional = new(@"^ФТД", RegexOptions.Compiled);
        HashSet<string> m_competenceList = null;
        EducationalWork m_eduWork = null;
        Department m_department = null;

        string m_competences = null;
        EDisciplineType? m_type = null;
        int m_startSemIdx = -1;
        int m_lastSemIdx = -1;

        /// <summary>
        /// Наименование дисциплины [ключ]
        /// </summary>
        public string Name { get; set; }
        public string Key { get => App.NormalizeName(Name); }
        /// <summary>
        /// Индекс дисциплины
        /// </summary>
        public string Index { get; set; }
        /// <summary>
        /// Тип дисциплины
        /// </summary>
        public EDisciplineType? Type {
            get {
                if (!string.IsNullOrEmpty(Index) && m_type == null) {
                    if (m_regexTestTypeOptional.IsMatch(Index)) {
                        m_type = EDisciplineType.Optional;
                    }
                    else if (m_regexTestTypeVariable.IsMatch(Index)) {
                        m_type = EDisciplineType.Variable;
                    }
                    else if (m_regexTestTypeByChoice.IsMatch(Index)) {
                        m_type = EDisciplineType.ByChoice;
                    }
                    else if (m_regexTestTypeRequired.IsMatch(Index)) {
                        m_type = EDisciplineType.Required;
                    }
                    else {
                        m_type = EDisciplineType.Unknown;
                    }
                }
                return m_type;
            }
        }
        /// <summary>
        /// Описание типа дисциплины [поле для РПД]
        /// </summary>
        public string TypeDescription {
            get {
                var text = "";
                var degree = Curriculum?.Degree == Enums.EDegree.Master ? "магистратуры" : "бакалавриата";

                if (Type == EDisciplineType.Required) {
                    text = $"обязательным дисциплинам программы {degree}";
                }
                else if (Type == EDisciplineType.Variable) {
                    text = $"обязательным дисциплинам части, формируемой участниками образовательных отношений программы {degree}";
                }
                else if (Type == EDisciplineType.ByChoice) {
                    text = $"дисциплинам по выбору части, формируемой участниками образовательных отношений программы {degree}";
                }
                else if (Type == EDisciplineType.Optional) {
                    text = $"программе {degree}, формируемой участниками образовательных отношений";
                }
                return text;
            }
        }

        /// <summary>
        /// Доп. описание типа дисциплины [поле для РПД]
        /// </summary>
        public string TypeDescription2 {
            get {
                var text = "";
                if (Type == EDisciplineType.Optional) {
                    text = " и относится к факультативной дисциплине";
                }
                
                return text;
            }
        }

        /// <summary>
        /// Тип контроля в регистре: зачет, экзамен
        /// </summary>
        public string ControlType { get => EducationalWork?.ControlFormForScreen.ToLower() ?? ""; }

        /// <summary>
        /// Название закрепленной кафедры
        /// </summary>
        public string DepartmentName { get; set; }
        /// <summary>
        /// Код закрепленной кафедры
        /// </summary>
        public string DepartmentCode { get; set; }
        /// <summary>
        /// Описание кафедры
        /// </summary>
        public Department Department { 
            get {
                if (m_department == null && !string.IsNullOrEmpty(DepartmentCode)) {
                    App.Config.Departments.TryGetValue(DepartmentCode, out m_department);
                }
                return m_department;
            }
        }
        /// <summary>
        /// Нормализованные компетенции (коды)
        /// </summary>
        public HashSet<string> CompetenceList {
            get {
                if (m_competenceList == null && !string.IsNullOrEmpty(Competences)) {
                    var competenceItems = Competences.Split(';', StringSplitOptions.TrimEntries);
                    m_competenceList = competenceItems.Select(x => string.Join("", x.Split(' ', StringSplitOptions.TrimEntries)).ToUpper()).ToHashSet();
                }
                return m_competenceList;
            }
        }
        /// <summary>
        /// Компетенции строкой (берутся из xlsx)
        /// </summary>
        public string Competences { get; set; }
        /// <summary>
        /// Формы контроля
        /// </summary>
        //public List<EControlForm> ControlForms { get; set; }
        public int? ControlFormExamHours { get; set; }
        public int? ControlFormTestHours { get; set; }
        public int? ControlFormTestWithAGradeHours { get; set; }
        public int? ControlFormControlWorkHours { get; set; }
        /// <summary>
        /// Итого акад.часов: по плану
        /// </summary>
        public int? TotalByPlanHours { get; set; }
        /// <summary>
        /// Итого акад.часов: Конт. раб.
        /// </summary>
        public int? TotalContactWorkHours { get; set; }
        /// <summary>
        /// Итого акад.часов: СР
        /// </summary>
        public int? TotalSelfStudyHours { get; set; }
        /// <summary>
        /// Итого акад.часов: Конт роль
        /// </summary>
        public int? TotalControlHours { get; set; }
        /// <summary>
        /// Описание учебных работ по семестрам
        /// </summary>
        public EducationalWork[] Semesters { get; set; } = new EducationalWork[SEMESTER_COUNT];

        /// <summary>
        /// Кол-во зачетных единиц (1 ЗЕ = 36 часов)
        /// </summary>
        public int? TestUnits { get => TotalByPlanHours / 36; }

        /// <summary>
        /// Выявленные ошибки
        /// </summary>
        public List<string> Errors { get; set; }
        /// <summary>
        /// Доп. ошибки (исп. при проверке по загруженной матрице компетенций)
        /// </summary>
        public List<string> ExtraErrors { get; set; }
        /// <summary>
        /// Родительский УП
        /// </summary>
        public Curriculum Curriculum { get; set; }

        public CurriculumDiscipline() {
            Errors = [];
            ExtraErrors = [];

            for (var i = 0; i < Semesters.Length; i++) {
                Semesters[i] = new();
            }
        }

        /// <summary>
        /// Объект описания учебной работы
        /// </summary>
        public EducationalWork EducationalWork {
            get {
                m_eduWork ??= new() {
                    ContactWorkHours = TotalContactWorkHours,
                    ControlForm = EControlForm.Unknown,
                    ControlHours = TotalControlHours,
                    LabHours = Semesters.Sum(s => s.LabHours ?? 0),
                    LectureHours = Semesters.Sum(s => s.LectureHours ?? 0),
                    PracticalHours = Semesters.Sum(s => s.PracticalHours ?? 0),
                    SelfStudyHours = TotalSelfStudyHours,
                    TotalHours = TotalByPlanHours
                };
                if (ControlFormExamHours.HasValue && ControlFormExamHours.Value > 0) {
                    m_eduWork.ControlForm = EControlForm.Exam;
                }
                else if (ControlFormTestHours.HasValue && ControlFormTestHours.Value > 0) {
                    m_eduWork.ControlForm = EControlForm.Test;
                }
                else if (ControlFormTestWithAGradeHours.HasValue && ControlFormTestWithAGradeHours.Value > 0) {
                    m_eduWork.ControlForm = EControlForm.TestWithAGrade;
                }
                else if (ControlFormControlWorkHours.HasValue && ControlFormControlWorkHours.Value > 0) {
                    m_eduWork.ControlForm = EControlForm.CourseWork;
                }
                
                return m_eduWork;
            }
        }

        /// <summary>
        /// Индекс стартового семестра
        /// </summary>
        public int StartSemesterIdx {
            get {
                if (m_startSemIdx < 0) {
                    var sem = this.Semesters.FirstOrDefault(s => s.TotalHours.HasValue && s.TotalHours > 0);
                    if (sem != null) {
                        m_startSemIdx = Array.IndexOf(this.Semesters, sem);
                    }
                }
                return m_startSemIdx;
            }
        }

        /// <summary>
        /// Индекс последнего семестра
        /// </summary>
        public int LastSemesterIdx {
            get {
                if (m_lastSemIdx < 0) {
                    var sem = this.Semesters.LastOrDefault(s => s.TotalHours.HasValue && s.TotalHours > 0);
                    if (sem != null) {
                        m_lastSemIdx = Array.IndexOf(this.Semesters, sem);
                    }
                }
                return m_lastSemIdx;
            }
        }

        /// <summary>
        /// Проверка дисциплины
        /// </summary>
        /// <param name="curriculum"></param>
        internal bool Check(Curriculum curriculum, int rowIdx) {
            var result = false;

            if (string.IsNullOrEmpty(Name)) {
                return result;
            }

            if (Type == EDisciplineType.Unknown) {
                var err = $"Не удалось определить тип дисциплины [{Name}] по индексу [{Index}] (строка {rowIdx})";
                curriculum.Errors.Add(err);
                Errors.Add(err);
            }
            //проверка итоговых час ов
            var total = (TotalContactWorkHours ?? 0) + (TotalControlHours ?? 0) + (TotalSelfStudyHours ?? 0);
            if (TotalByPlanHours != total) {
                var err = $"Дисциплина [{Name}] (строка {rowIdx}): обнаружено неверное значение в поле [По плану]";
                curriculum.Errors.Add(err);
                Errors.Add(err);
            }
            //проверка часов семестров
            var allSemTotal = 0;
            var allSemContact = 0;
            for (var i = 0; i < Semesters.Length; i++) {
                var sem = Semesters[i];
                var semTotal = (sem.LectureHours ?? 0) + (sem.LabHours ?? 0) + (sem.SelfStudyHours ?? 0) + (sem.PracticalHours ?? 0) + (sem.ControlHours ?? 0);
                if (sem.TotalHours != semTotal) {
                    var err = $"Дисциплина [{Name}] (строка {rowIdx}): обнаружено неверное значение в поле [Итого] Семестра {i + 1}";
                    curriculum.Errors.Add(err);
                    Errors.Add(err);
                }
                allSemTotal += semTotal;
                var semContact = (sem.LectureHours ?? 0) + (sem.LabHours ?? 0) + (sem.PracticalHours ?? 0);
                allSemContact += semContact;
            }
            if (TotalByPlanHours != allSemTotal) {
                var err = $"Дисциплина [{Name}] (строка {rowIdx}): сумма итогового времени по семестрам не сходится с полем [Итого]";
                curriculum.Errors.Add(err);
                Errors.Add(err);
            }
            if (TotalContactWorkHours != allSemContact) {
                var err = $"Дисциплина [{Name}] (строка {rowIdx}): сумма контактного времени по семестрам не сходится с полем [Конт. раб.]";
                curriculum.Errors.Add(err);
                Errors.Add(err);
            }
            //if (TotalContactWorkHours != prac)
            //пропуска строку с названием модуля - у нее не заполнена ячейка кафедры
            if (!string.IsNullOrEmpty(DepartmentName)) {
                result = true;
            }

            return result;
        }

        /// <summary>
        /// Проверка компетенций по указанной матрице
        /// </summary>
        /// <returns></returns>
        internal bool CheckCompetences(CompetenceMatrix matrix) {
            ExtraErrors = [];

            if (matrix == null) {
                ExtraErrors.Add("Матрица компетенций не загружена. Проверить компетенции не удалось.");
                return false;
            }

            var achiCodeList = matrix.GetAllAchievementCodes();
            if (CompetenceList == null || CompetenceList.Count == 0) {
                ExtraErrors.Add($"Список компетенций не определён.");
            }
            else {
                foreach (var achiCode in CompetenceList) {
                    if (!achiCodeList.Contains(achiCode)) {
                        ExtraErrors.Add($"В загруженной матрице компетенций не найден индикатор достижений [{achiCode}].");
                    }
                }
            }

            return ExtraErrors.Count == 0;
        }
    }
}
