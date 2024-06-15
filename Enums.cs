using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    internal class Enums {
        /// <summary>
        /// Форма обучения
        /// </summary>
        public enum EFormOfStudy {
            [Description("Очная")]
            FullTime,               //очная
            [Description("Заочная")]
            PartTime,               //заочная
            [Description("Очно-заочная")]
            MixedTime,              //очно-заочная
            [Description("НЕИЗВЕСТНО")]
            Unknown
        }

        /// <summary>
        /// Квалификация
        /// </summary>
        public enum EDegree {
            [Description("магистр")]
            Master,
            [Description("бакалавр")]
            Bachelor,
            [Description("НЕИЗВЕСТНО")]
            Unknown
        }

        /// <summary>
        /// Тип оценочного средства
        /// </summary>
        public enum EEvaluationTool {
            [Description("Опрос")]
            Survey,         //опрос
            [Description("Эссе")]
            Essay,          //реферат
            [Description("Реферат")]
            Paper,
            [Description("Тестирование")]
            Testing,        //тестирование
            [Description("Контрольная работа")]
            ControlWork,    //контрольная работа
            [Description("Доклад")]
            Presentation
            //PracticalWork   //практическая работа
        }

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
    }
}
