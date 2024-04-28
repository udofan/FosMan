using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
    /// Описание дисциплины
    /// </summary>
    internal class CurriculumDiscipline {
        /// <summary>
        /// Наименование дисциплины [ключ]
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Индекс дисциплины
        /// </summary>
        public string Index { get; set; }
        /// <summary>
        /// Закрепленная кафедра
        /// </summary>
        public string Department { get; set; }
        /// <summary>
        /// Компетенции (коды)
        /// </summary>
        public List<string> CompetenceList { get; set; }
        /// <summary>
        /// Формы контроля
        /// </summary>
        public List<EControlForm> ControlForms { get; set; }
    }
}
