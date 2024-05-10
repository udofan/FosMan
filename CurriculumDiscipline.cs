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
        ControlWork,    //контрольная работа
        [Description("НЕИЗВЕСТНО")]
        Unknown         //определить НЕ УДАЛОСЬ
    }

    /// <summary>
    /// Тип дисциплины
    /// </summary>
    public enum EDisciplineType {
        [Description("Обязательная")]
        Required,       //обязательная
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
    internal class CurriculumDiscipline {
        public const int SEMESTER_COUNT = 10;

        static TypeAccessor m_typeAccessor = TypeAccessor.Create(typeof(CurriculumDiscipline));
        static Dictionary<Type, TypeAccessor> m_extraTypeAccessors = new();
        static Regex m_regexTestTypeRequired = new(@"^[^\.]+\.О\.", RegexOptions.Compiled);
        static Regex m_regexTestTypeByChoice = new(@"^[^\.]+\.В\.", RegexOptions.Compiled);
        static Regex m_regexTestTypeOptional = new(@"^ФТД", RegexOptions.Compiled);
        HashSet<string> m_competenceList = null;
        EducationalWork m_eduWork = null;

        string m_competences = null;
        EDisciplineType? m_type = null;

        /// <summary>
        /// Наименование дисциплины [ключ]
        /// </summary>
        public string Name { get; set; }
        public string Key { get => Name?.ToUpper(); }
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
                if (Type == EDisciplineType.Required) {
                    text = "обязательным дисциплинам";
                }
                else if (Type == EDisciplineType.ByChoice) {
                    text = "дисциплинам по выбору части, формируемой участниками образовательных отношений";
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
        /// Закрепленная кафедра
        /// </summary>
        public string Department { get; set; }
        /// <summary>
        /// Код закрепленной кафедры
        /// </summary>
        public string DepartmentCode { get; set; }
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
                if (ControlFormExamHours.HasValue && ControlFormExamHours.Value > 0) m_eduWork.ControlForm = EControlForm.Exam;
                if (ControlFormTestHours.HasValue && ControlFormTestHours.Value > 0) m_eduWork.ControlForm = EControlForm.Test;
                if (ControlFormTestWithAGradeHours.HasValue && ControlFormTestWithAGradeHours.Value > 0) m_eduWork.ControlForm = EControlForm.TestWithAGrade;
                if (ControlFormControlWorkHours.HasValue && ControlFormControlWorkHours.Value > 0) m_eduWork.ControlForm = EControlForm.ControlWork;
                
                return m_eduWork;
            }
        }

        /// <summary>
        /// Установить значение указанному свойству
        /// </summary>
        /// <param name="targetProperty">целевое свойство</param>
        /// <param name="targetType">тип целевого свойства</param>
        /// <param name="cellValue">значение (будет приведено к типу targetType)</param>
        /// <param name="index">индекс списка, если целевое свойство List(object)</param>
        /// <param name="subProperty">свойство object'а из списка</param>
        internal void SetProperty(string targetProperty, Type targetType, string cellValue, int index = -1, string subProperty = null) {
            object value = null;

            cellValue = cellValue?.Trim();

            if (targetType == typeof(int)) {
                if (int.TryParse(cellValue, out var intValue)) {
                    value = intValue;
                }
            }
            else if (targetType == typeof(string)) {
                value = cellValue;
            }

            if (index >= 0 && !string.IsNullOrEmpty(subProperty)) {
                var arrayObj = m_typeAccessor[this, targetProperty] as object[];
                if (arrayObj != null) {
                    var extraType = arrayObj.FirstOrDefault()?.GetType();
                    if (extraType != null) {
                        if (!m_extraTypeAccessors.TryGetValue(extraType, out var extratypeAccessor)) {
                            extratypeAccessor = TypeAccessor.Create(extraType);
                            m_extraTypeAccessors[extraType] = extratypeAccessor;
                        }
                        extratypeAccessor[arrayObj[index], subProperty] = value;
                    }
                }
            }
            else {
                m_typeAccessor[this, targetProperty] = value;
            }
        }

        /// <summary>
        /// Проверка дисциплины
        /// </summary>
        /// <param name="curriculum"></param>
        internal bool Check(Curriculum curriculum, int rowIdx) {
            if (string.IsNullOrEmpty(Name)) {
                return false;
            }

            if (Type == EDisciplineType.Unknown) {
                var err = $"Не удалось определить тип дисциплины [{Name}] по индексу [{Index}] (строка {rowIdx})";
                curriculum.Errors.Add(err);
                Errors.Add(err);
            }
            //проверка итоговых часов
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

            return true;
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
            foreach (var achiCode in CompetenceList) { 
                if (!achiCodeList.Contains(achiCode)) {
                    ExtraErrors.Add($"В загруженной матрице компетенций не найден индикатор достижений [{achiCode}].");
                }
            }

            return ExtraErrors.Count == 0;
        }

        /// <summary>
        /// Получить значение свойства по имени
        /// </summary>
        /// <param name="propName"></param>
        /// <returns></returns>
        public object GetProperty(string propName) {
            object value = null;
            try {
                value = m_typeAccessor[this, propName];
            }
            catch (Exception ex) {
            }

            return value;
        }
    }
}
