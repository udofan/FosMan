using FastMember;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    /// <summary>
    /// Описание учебной работы
    /// </summary>
    internal class EducationalWork {
        /// <summary>
        /// Аксессор для типа для прямой работы со свойствами через их имя
        /// </summary>
        static public TypeAccessor TypeAccessor { get; } = TypeAccessor.Create(typeof(EducationalWork));

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
        /// <summary>
        /// Значение ControlForm для экрана
        /// </summary>
        public string ControlFormForScreen { get => ControlForm.GetDescription(); }
        /// <summary>
        /// Таблица учебного времени с темами
        /// </summary>
        public Table Table { get; set; }

        /// <summary>
        /// Получить значение свойства по имени
        /// </summary>
        /// <param name="propName"></param>
        /// <returns></returns>
        public object GetProperty(string propName) {
            object value = null;
            try {
                value = TypeAccessor[this, propName];
            }
            catch (Exception ex) {
            }

            return value;
        }
    }
}
