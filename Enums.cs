using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    public class EvaluationToolAttribute : Attribute {
        /// <summary>
        /// Флаг однократного применение оценочного средства
        /// </summary>
        public bool SingleUse { get; set; }
        /// <summary>
        /// Короткое экранное название
        /// </summary>
        public string ShortDescription { get; set; }

        public EvaluationToolAttribute(string shortDescription, bool singleUse = false) {
            ShortDescription = shortDescription;
            SingleUse = singleUse;
        }
    }

    public static class Enums {
        /// <summary>
        /// Форма обучения
        /// </summary>
        public enum EFormOfStudy {
            [Description("Очная")]
            [EvaluationTool("О")]
            FullTime = 0,               //очная
            [Description("Очно-заочная")]
            [EvaluationTool("ОЗ")]
            MixedTime = 1,              //очно-заочная
            [Description("Заочная")]
            [EvaluationTool("З")]
            PartTime = 2,               //заочная
            [Description("НЕИЗВЕСТНО")]
            [EvaluationTool("?")]
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
            [EvaluationTool("О")]
            Survey,
            [Description("Эссе")]
            [EvaluationTool("Э")]
            Essay,
            [Description("Реферат")]
            [EvaluationTool("Р")]
            Paper,
            [Description("Тестирование")]
            [EvaluationTool("Т")]
            Testing,
            [Description("Контрольная работа")]
            [EvaluationTool("К")]
            ControlWork,
            [Description("Доклад")]
            [EvaluationTool("Д")]
            Presentation,
            [Description("Практическая работа")]
            [EvaluationTool("ПР")]
            PracticalWork,
            [Description("Мини-кейсы")]
            [EvaluationTool("МК")]
            MiniCases,
            [Description("Деловая игра")]
            [EvaluationTool("ДИ")]
            BusinessGame,
            [Description("Курсовая работа")]
            [EvaluationTool("УР", true)]
            CourseWork
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
            /// Фикс: Время (всего, лекции, практика, самоподготовка)
            /// </summary>
            Time = 0b0000_0001,
            /// <summary>
            /// Фиск: Оценочные средства
            /// </summary>
            EvalTools = 0b0000_0010,
            /// <summary>
            /// Фиск: Результаты компетенций
            /// </summary>
            CompetenceResults = 0b0000_0100,
            /// <summary>
            /// Фиск: Полное перестроение
            /// </summary>
            FullRecreate = 0b0000_1000,
            /// <summary>
            /// Опция: Брать оценочные средства из ФОС (по-возможности)
            /// </summary>
            TakeEvalToolsFromFos = 0b0001_0000,
            /// <summary>
            /// Все виды фиксов
            /// </summary>
            All = 0b0000_1111
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
            [Description("Не удалось определить цель дисциплины")]
            RpdMissingPurpose,
            [Description("Не удалось определить описание дисциплины")]
            RpdMissingDescription,
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
            CompetenceMatrixResultCodeAchievementMismatch,
            [Description("Не определены StopMarkers для правила с типом Multiline")]
            ParseRuleMissingStopMarkers,
            [Description("Некорректный [inlineGroupIdx] для элемента значения StartMarkers")]
            ParseRuleWrongInlineGroupIdxInStartMarkers,
            [Description("Не найдена РПД для дисциплины")]
            GenAbstractsMissingRpd,
            [Description("Не удалось получить значение свойства")]
            GenAbstractsMissingProperty,
            [Description("Простая ошибка")]
            Simple
        }

        /// <summary>
        /// Типы файлов для режима коррекции
        /// </summary>
        public enum EFileType {
            /// <summary>
            /// Автоматическое определение
            /// </summary>
            Auto,
            /// <summary>
            /// Неизвестный тип
            /// </summary>
            [Description("Не определён")]
            Unknown,
            /// <summary>
            /// РПД
            /// </summary>
            [Description("РПД")]
            Rpd,
            /// <summary>
            /// ФОС
            /// </summary>
            [Description("ФОС")]
            Fos
        }

        static Dictionary<string, EEvaluationTool> m_evalToolDic = null;

        /// <summary>
        /// Словарь видов оценочных средств
        /// </summary>
        public static Dictionary<string, EEvaluationTool> EvalToolDic {
            get {
                m_evalToolDic ??= Enum.GetValues(typeof(EEvaluationTool)).Cast<EEvaluationTool>().ToDictionary(x => x.GetDescription().ToUpper(), x => x);

                return m_evalToolDic;
            }
        }
    }
}
