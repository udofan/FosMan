using FastMember;
using System;
using System.CodeDom;
using System.Collections.Generic;
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
        Exam,           //экзамен
        Test,           //зачет
        TestWithAGrade, //зачет с оценкой
        ControlWork     //контрольная работа
    }

    /// <summary>
    /// Тип дисциплины
    /// </summary>
    public enum EDisciplineType {
        Required,       //обязательная
        ByChoice,       //по выбору
        Optional,       //факультативная
        Unknown         //опередить НЕ УДАЛОСЬ
    }

    /// <summary>
    /// Описание дисциплины
    /// </summary>
    internal class CurriculumDiscipline {
        static TypeAccessor m_typeAccesor = TypeAccessor.Create(typeof(CurriculumDiscipline));
        static Dictionary<Type, TypeAccessor> m_extraTypeAccessors = new();
        static Regex m_regexTestTypeRequired = new(@"^[^\.]+\.О\.", RegexOptions.Compiled);
        static Regex m_regexTestTypeByChoice = new(@"^[^\.]+\.В\.", RegexOptions.Compiled);
        static Regex m_regexTestTypeOptional = new(@"^ФТД", RegexOptions.Compiled);

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
        /// Закрепленная кафедра
        /// </summary>
        public string Department { get; set; }
        /// <summary>
        /// Код закрепленной кафедры
        /// </summary>
        public string DepartmentCode { get; set; }
        /// <summary>
        /// Компетенции (коды)
        /// </summary>
        public List<string> CompetenceList { get; set; }

        public string Competences {
            get => m_competences;
            set {
                if (!string.IsNullOrEmpty(value)) {
                    var competenceItems = value.Split(';', StringSplitOptions.TrimEntries);
                    CompetenceList = [];
                    CompetenceList.AddRange(competenceItems.Select(x => string.Join("", x.Split(' '))));
                }
            }
        }
        /// <summary>
        /// Формы контроля
        /// </summary>
        public List<EControlForm> ControlForms { get; set; }
        public int? ControlFormExamHours { get; set; } = 0;
        public int? ControlFormTestHours { get; set; } = 0;
        public int? ControlFormTestWithAGradeHours { get; set; } = 0;
        public int? ControlFormControlWorkHours { get; set; } = 0;
        /// <summary>
        /// Итого акад.часов: по плану
        /// </summary>
        public int? TotalByPlanHours { get; set; } = 0;
        /// <summary>
        /// Итого акад.часов: Конт. раб.
        /// </summary>
        public int? TotalContactWorkHours { get; set; } = 0;
        /// <summary>
        /// Итого акад.часов: СР
        /// </summary>
        public int? TotalSelfStudyHours { get; set; } = 0;
        /// <summary>
        /// Итого акад.часов: Конт роль
        /// </summary>
        public int? TotalControlHours { get; set; } = 0;
        /// <summary>
        /// Описание семестров
        /// </summary>
        public CurriculumDisciplineSemester[] Semesters { get; set; } = new CurriculumDisciplineSemester[8];

        /// <summary>
        /// Выявленные ошибки
        /// </summary>
        public List<string> Errors { get; set; }

        public CurriculumDiscipline() {
            Errors = new();

            for (var i = 0; i < Semesters.Length; i++) {
                Semesters[i] = new();
            }
        }

        /// <summary>
        /// Установить значение указанному свойству
        /// </summary>
        /// <param name="targetProperty"></param>
        /// <param name="targetType"></param>
        /// <param name="cellValue"></param>
        /// <exception cref="NotImplementedException"></exception>
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
                var arrayObj = m_typeAccesor[this, targetProperty] as object[];
                if (arrayObj != null) {
                    var extraType = arrayObj.FirstOrDefault()?.GetType();
                    if (extraType != null) {
                        if (!m_extraTypeAccessors.TryGetValue(extraType, out var extraTypeAccesor)) {
                            extraTypeAccesor = TypeAccessor.Create(extraType);
                            m_extraTypeAccessors[extraType] = extraTypeAccesor;
                        }
                        extraTypeAccesor[arrayObj[index], subProperty] = value;
                    }
                }
            }
            else {
                m_typeAccesor[this, targetProperty] = value;
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

            return true;
        }
    }
}
