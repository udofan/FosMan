using FastMember;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Xceed.Document.NET;
using static FosMan.Enums;

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
        [JsonInclude]
        public int? TotalHours { get; set; } = 0;
        /// <summary>
        /// Контактная (аудиторная) работы
        /// </summary>
        [JsonInclude]
        public int? ContactWorkHours { get; set; } = 0;
        /// <summary>
        /// Лекции
        /// </summary>
        [JsonInclude]
        public int? LectureHours { get; set; } = 0;
        /// <summary>
        /// Лабораторные работы
        /// </summary>
        [JsonInclude]
        public int? LabHours { get; set; } = 0;
        /// <summary>
        /// Практические занятия
        /// </summary>
        [JsonInclude] 
        public int? PracticalHours { get; set; } = 0;
        /// <summary>
        /// Самостоятельная работы
        /// </summary>
        [JsonInclude] 
        public int? SelfStudyHours { get; set; } = 0;
        /// <summary>
        /// Контроль
        /// </summary>
        [JsonInclude] 
        public int? ControlHours { get; set; } = 0;
        /// <summary>
        /// Форма итогового контроля
        /// </summary>
        [JsonInclude] 
        public EControlForm ControlForm { get; set; } = EControlForm.Unknown;
        /// <summary>
        /// Значение ControlForm для экрана
        /// </summary>
        [JsonInclude] 
        public string ControlFormForScreen { get => ControlForm.GetDescription(); }
        /// <summary>
        /// Таблица учебного времени с темами
        /// </summary>
        [JsonIgnore]
        internal Table Table { get; set; }
        /// <summary>
        /// Модули обучения
        /// </summary>
        [JsonInclude]
        public List<StudyModule> Modules { get; set; }
        [JsonIgnore]
        public int TableStartNumCol { get; set; } = -1;
        [JsonIgnore]
        public int TableTopicStartRow { get; set; } = -1;
        [JsonIgnore]
        public int TableTopicLastRow { get; set; } = -1;
        /// <summary>
        /// Таблица: кол-во колонок
        /// </summary>
        [JsonIgnore]
        public int TableMaxColCount { get; set; } = -1;
        /// <summary>
        /// Таблица: в таблице есть колонка подитога по контактным работам
        /// </summary>
        [JsonIgnore]
        public bool TableHasContactTimeSubtotal { get; set; } = false;
        /// <summary>
        /// Таблица: номер ряда с контролем (зачет/экзамен)
        /// </summary>
        [JsonIgnore]
        public int TableControlRow { get; set; } = -1;
        [JsonIgnore]
        public int TableColTopic { get; set; } = -1;
        [JsonIgnore]
        public int TableColEvalTools { get; set; } = -1;
        [JsonIgnore]
        public int TableColCompetenceResults { get; set; } = -1;

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
