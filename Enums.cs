using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    public class Enums {
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
            Survey,
            [Description("Эссе")]
            Essay,
            [Description("Реферат")]
            Paper,
            [Description("Тестирование")]
            Testing,
            [Description("Контрольная работа")]
            ControlWork,
            [Description("Доклад")]
            Presentation,
            [Description("Практическая работа")]
            PracticalWork,
            [Description("Мини-кейсы")]
            MiniCases,
            [Description("Деловая игра")]
            BusinessGame
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

        /// <summary>
        /// Формат таблицы компетенций
        /// </summary>
        public enum ECompetenceMatrixFormat {
            Unknown,
            Rpd,
            Fos21,
            Fos22
        }

        /// <summary>
        /// Виды элементов в Сторе
        /// </summary>
        [Flags]
        public enum EStoreElements {
            Rpd = 0b0000_0001,
            Fos = 0b0000_0010,
            Curricula = 0b0000_0100,
            All = 0b1111_1111
        }
    }
}
