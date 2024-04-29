using FastMember;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

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
        public int ?ControlFormTestWithAGradeHours { get; set; } = 0;
        public int? ControlFormControlWorkHours { get; set; } = 0;

        /// <summary>
        /// Установить значение указанному свойству
        /// </summary>
        /// <param name="targetProperty"></param>
        /// <param name="targetType"></param>
        /// <param name="cellValue"></param>
        /// <exception cref="NotImplementedException"></exception>
        internal void SetProperty(string targetProperty, Type targetType, string cellValue) {
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

            m_typeAccesor[this, targetProperty] = value;
        }
    }
}
