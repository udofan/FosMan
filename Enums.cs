﻿using System;
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

        /// <summary>
        /// Типы фиксов таблицы содержания дисциплин (РПД)
        /// </summary>
        [Flags]
        public enum EEduWorkFixType {
            /// <summary>
            /// Значение не определено
            /// </summary>
            Undefined = 0,
            /// <summary>
            /// Время (всего, лекции, практика, самоподготовка)
            /// </summary>
            Time = 0b0000_0001,
            /// <summary>
            /// Оценочные средства
            /// </summary>
            EvalTools = 0b0000_0010,
            /// <summary>
            /// Результаты компетенций
            /// </summary>
            CompetenceResults = 0b0000_0100,
            /// <summary>
            /// Все виды фиксов
            /// </summary>
            All = 0b1111_1111
        }

        /// <summary>
        /// Типы ошибок
        /// </summary>
        public enum EErrorType {
            [Description("Ошибка")]
            Exception,
            [Description("Не задан тип оценочного средства")]
            RpdMissingEvalTool,
            [Description("Не определён тип оценочного средства")]
            RpdUnknownEvalTool,
            [Description("Не найдена информация по учебным работам")]
            RpdMissingEduWork,
            [Description("Не найдена матрица компетенций")]
            RpdMissingCompetenceMatrix,
            [Description("Не удалось обнаружить содержание разделов и тем")]
            RpdMissingTOC,
            [Description("Не удалось обнаружить список вопросов к зачету/экзамену")]
            RpdMissingQuestionList,
            [Description("Учебные работы обнаружены не по всем формам обучения")]
            RpdNotFullEduWorks,
            [Description("Не удалось обнаружить список основной литературы")]
            RpdMissingReferencesBase,
            [Description("Не удалось обнаружить список дополнительной литературы")]
            RpdMissingReferencesExtra,
            [Description("Не удалось обнаружить таблицу учебных работ")]
            RpdMissingEduWorkTable,
            [Description("Не удалось определить название кафедры")]
            RpdMissingDepartment,
            [Description("Не удалось определить профиль")]
            RpdMissingProfile,
            [Description("Не удалось определить год программы")]
            RpdMissingYear,
            [Description("Не удалось определить шифр направления подготовки")]
            RpdMissingDirectionCode,
            [Description("Не удалось определить наименование направления подготовки")]
            RpdMissingDirectionName,
            [Description("Не удалось определить название дисциплины")]
            RpdMissingDisciplineName,
            [Description("В документе не найдено таблиц")]
            CompetenceMatrixNoTables,
            [Description("Формат 2.2: в матрице не удалось найти компетенцию")]
            CompetenceMatrixMissingItem,
            [Description("Не удалось распарсить текст компетенции")]
            CompetenceMatrixItemParseError,
            [Description("Формат 2.2: в матрице не удалось найти достижение")]
            CompetenceMatrixMissingAchievement,
            [Description("Не удалось распарсить текст индикатора достижения")]
            CompetenceMatrixAchievementParseError,
            [Description("Не удалось распарсить текст результата")]
            CompetenceMatrixResultParseError,
            [Description("Не указаны коды результата")]
            CompetenceMatrixMissingResult,
            [Description("Не удалось определить номер семестра")]
            CompetenceMatrixSemesterParseError,
            [Description("Не указан семестр")]
            CompetenceMatrixMissingSemester,
            [Description("Список компетенций не определён")]
            CompetenceMatrixIsEmpty,
            [Description("Список индикаторов достижения компетенции не определён")]
            CompetenceMatrixItemMissingAchievements,
            [Description("Индикатор достижения компетенции не определён")]
            CompetenceMatrixItemMissingAchievementCode,
            [Description("Индикатор достижения не соответствует компетенции")]
            CompetenceMatrixAchievementCodeMismatch,
            [Description("Не определён результат индикатора достижения компетенции")]
            CompetenceMatrixMissingAchievementResult,
            [Description("Результат индикатора достижения не соответствует компетенции")]
            CompetenceMatrixResultCodeItemMismatch,
            [Description("Результат индикатора достижения не соответствует индикатору")]
            CompetenceMatrixResultCodeAchievementMismatch
        }
    }
}
