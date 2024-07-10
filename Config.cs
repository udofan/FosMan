using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Windows.Resources;
using static FosMan.Enums;

namespace FosMan {
    internal class Config {
        bool m_rpdFixEduWorkTablesFullRecreate = false;

        [JsonInclude]
        public string YaOAuthToken { get; set; }
        [JsonInclude]
        public string YaGptIamToken { get; set; }
        [JsonInclude]
        public DateTimeOffset YaGptIamTokenExpiresAt { get; set; }
        /// <summary>
        /// Файл матрицы компетенций
        /// </summary>
        [JsonInclude]
        public string CompetenceMatrixFileName { get; set; }
        /// <summary>
        /// Автозагрузка матрицы при запуске приложения
        /// </summary>
        [JsonInclude]
        public bool CompetenceMatrixAutoload { get; set; } = true;
        /// <summary>
        /// Последняя выбранная директория для добавления файлов
        /// </summary>
        [JsonInclude]
        public string RpdLastAddDir { get; set; } = Directory.GetCurrentDirectory();
        /// <summary>
        /// Последняя выбранная локация РПД-файлов
        /// </summary>
        [JsonInclude]
        public string RpdLastLocation { get; set; }
        /// <summary>
        /// Флаг сохранения списка загруженных РПД
        /// </summary>
        [JsonInclude]
        public bool StoreRpdList { get; set; } = true;
        /// <summary>
        /// Сохраненный список файлов РПД
        /// </summary>
        [JsonInclude]
        public List<string> RpdList { get; set; } = [];
        /// <summary>
        /// Фикс таблицы компетенций
        /// </summary>
        [JsonInclude]
        public bool RpdFixTableOfCompetences { get; set; } = true;
        /// <summary>
        /// Фикс таблицы учебных работ
        /// </summary>
        [JsonInclude]
        public bool RpdFixTableOfEduWorks { get; set; } = true;
        /// <summary>
        /// Применять настройку "найти и заменить" для режима исправлений РПД
        /// </summary>
        [JsonInclude]
        public bool RpdFixFindAndReplace { get; set; } = false;
        /// <summary>
        /// Элементы списка "найти и заменить" для режима исправлений РПД
        /// </summary>
        [JsonInclude]
        public List<FindAndReplaceItem> RpdFixFindAndReplaceItems { get; set; } = [];
        /// <summary>
        /// Целевая директория для генерируемых РПД
        /// </summary>
        [JsonInclude]
        public string RpdFixTargetDir { get; set; }
        [JsonInclude]
        public bool RpdFixByTemplate { get; set; }
        [JsonInclude]
        public string RpdFixTemplateFileName { get; set; }
        /// <summary>
        /// Заполнение таблиц учебных работ для форм обучения: распределение времени по темам
        /// </summary>
        [JsonInclude]
        public bool RpdFixEduWorkTablesFixTime { get; set; }
        /// <summary>
        /// Заполнение таблиц учебных работ для форм обучения: расстановка оценочных средств
        /// </summary>
        [JsonInclude]
        public bool RpdFixEduWorkTablesFixEvalTools { get; set; }
        /// <summary>
        /// Заполнение таблиц учебных работ для форм обучения: расстановка результатов компетенций случайным образом
        /// </summary>
        [JsonInclude]
        public bool RpdFixEduWorkTablesFixCompetenceCodes { get; set; }
        /// <summary>
        /// Максимальное значение результатов компетенций при их автоматической расстановке
        /// </summary>
        [JsonInclude]
        public decimal RpdFixMaxCompetenceResultsCount { get; set; } = 3;
        /// <summary>
        /// Флаг полного перестроения таблицы содержания дисциплины
        /// </summary>
        [JsonInclude]
        public bool RpdFixEduWorkTablesFullRecreate {
            get => m_rpdFixEduWorkTablesFullRecreate;
            set {
                m_rpdFixEduWorkTablesFullRecreate = value;
                if (value) {
                    RpdFixEduWorkTablesFixCompetenceCodes = true;
                    RpdFixEduWorkTablesFixEvalTools = true;
                    RpdFixEduWorkTablesFixTime = true;
                }
            }
        }
        /// <summary>
        /// Список оценочных средств, выдаваемых первым темам при их расстановке
        /// </summary>
        [JsonInclude]
        public List<EEvaluationTool> RpdFixEduWorkTablesEvalTools1stStageItems { get; set; } = [EEvaluationTool.Survey, EEvaluationTool.Testing];
        /// <summary>
        /// Полный список оценочных средств, выдаваемых темам при их расстановке
        /// </summary>
        [JsonInclude]
        public List<EEvaluationTool> RpdFixEduWorkTablesEvalTools2ndStageItems { get; set; } = [EEvaluationTool.Essay, EEvaluationTool.Paper, EEvaluationTool.Presentation, EEvaluationTool.ControlWork];
        /// <summary>
        /// Задать списки предшествующих и последующих дисциплин
        /// </summary>
        [JsonInclude]
        public bool RpdFixSetPrevAndNextDisciplines { get; set; }
        /// <summary>
        /// Убирать выделение цветом служебных областей
        /// </summary>
        [JsonInclude]
        public bool RpdFixRemoveColorSelections { get; set; }
        /// <summary>
        /// Последняя выбранная локация файлов УП
        /// </summary>
        [JsonInclude]
        public string CurriculumLastLocation { get; set; }
        /// <summary>
        /// Список свойств документа РПД
        /// </summary>
        [JsonInclude]
        public List<DocProperty> RpdFixDocPropertyList { get; set; }
        /// <summary>
        /// Коррекция РПД: брать оценочные средства из ФОС
        /// </summary>
        [JsonInclude]
        public bool RpdFixEduWorkTablesTakeEvalToolsFromFos { get; set; }
        /// <summary>
        /// Флаг сохранения списка загруженных УП
        /// </summary>
        [JsonInclude]
        public bool StoreCurriculumList { get; set; } = true;
        /// <summary>
        /// Сохраненный список файлов УП
        /// </summary>
        [JsonInclude]
        public List<string> CurriculumList { get; set; } = [];
        /// <summary>
        /// Шаблон для генерации РПД
        /// </summary>
        [JsonInclude]
        public string RpdGenTemplate { get; set; } = null;
        /// <summary>
        /// Шаблона для имени файла генерируемого РПД
        /// </summary>
        [JsonInclude]
        public string RpdGenFileNameTemplate { get; set; } = "РПД_{Index}_{Name}_2024.docx";
        /// <summary>
        /// Целевая директория для генерируемых РПД
        /// </summary>
        [JsonInclude]

        public string RpdGenTargetDir { get; set; }
        /// <summary>
        /// Последняя выбранная директория для добавления файлов ФОС
        /// </summary>
        [JsonInclude]
        public string FosLastAddDir { get; set; } = Directory.GetCurrentDirectory();
        /// <summary>
        /// Последняя выбранная локация ФОС-файлов
        /// </summary>
        [JsonInclude]
        public string FosLastLocation { get; set; }
        /// <summary>
        /// Флаг сохранения списка загруженных ФОС
        /// </summary>
        [JsonInclude]
        public bool StoreFosList { get; set; } = true;
        /// <summary>
        /// Сохраненный список файлов ФОС
        /// </summary>
        [JsonInclude]
        public List<string> FosList { get; set; } = [];
        /// <summary>
        /// Фикс таблицы компетенций #1
        /// </summary>
        [JsonInclude]
        public bool FosFixCompetenceTable1 { get; set; }
        /// <summary>
        /// Фикс таблицы компетенций #2
        /// </summary>
        [JsonInclude]
        public bool FosFixCompetenceTable2 { get; set; }
        /// <summary>
        /// Фикс таблицы паспорта
        /// </summary>
        [JsonInclude]
        public bool FosFixPassportTable { get; set; }
        /// <summary>
        /// Сброс цветовых выделений
        /// </summary>
        [JsonInclude]
        public bool FosFixResetSelection { get; set; }
        /// <summary>
        /// Коррекция кодов индикаторов компетенций в таблицах описания оценочных средств
        /// </summary>
        [JsonInclude]
        public bool FosFixCompetenceIndicators { get; set; }
        /// <summary>
        /// Сброс цветовых выделений
        /// </summary>
        [JsonInclude]
        public string FosFixTargetDir { get; set; }
        /// <summary>
        /// Элементы списка "найти и заменить" для режима исправлений ФОС
        /// </summary>
        [JsonInclude]
        public List<FindAndReplaceItem> FosFixFindAndReplaceItems { get; set; } = [];
        /// <summary>
        /// Список свойств документа ФОС
        /// </summary>
        [JsonInclude]
        public List<DocProperty> FosFixDocPropertyList { get; set; }
        /// <summary>
        /// Последняя директория, откуда добавлялись файлы в режим коррекции файлов
        /// </summary>
        [JsonInclude]
        public string FileFixerLastDirectory { get; set; }
        /// <summary>
        /// Список описаний кафедр
        /// </summary>
        [JsonInclude]
        public Dictionary<string, Department> Departments { get; set; }
        /// <summary>
        /// Элементы для парсинга заголовков excel-файлов УП
        /// </summary>
        [JsonInclude]
        public List<CurriculumDisciplineHeader> CurriculumDisciplineParseItems { get; set; }
    }
}
