using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Описание учебной работы
    /// </summary>
    internal class EducationalWork {
        /// <summary>
        /// Общая трудоемкость
        /// </summary>
        public int? TotalHours { get; set; } = 0;
        /// <summary>
        /// Контактная (аудиторная) работы
        /// </summary>
        public int? ContactWorkHours { get; set; } = 0;
        /// <summary>
        /// Лекции
        /// </summary>
        public int? LectureHours { get; set; } = 0;
        /// <summary>
        /// Лабораторные работы
        /// </summary>
        public int? LabHours { get; set; } = 0;
        /// <summary>
        /// Практические занятия
        /// </summary>
        public int? PracticalHours { get; set; } = 0;
        /// <summary>
        /// Самостоятельная работы
        /// </summary>
        public int? SelfStudyHours { get; set; } = 0;
        /// <summary>
        /// Контроль
        /// </summary>
        public int? ControlHours { get; set; } = 0;
        /// <summary>
        /// Форма итогового контроля
        /// </summary>
        public EControlForm ControlForm { get; set; } = EControlForm.Unknown;

    }
}
