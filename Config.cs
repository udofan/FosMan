using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace FosMan {
    internal class Config {
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
        /// Элементы списка "найти и заменить" для режима исправлений РПД
        /// </summary>
        [JsonInclude]
        public List<RpdFindAndReplaceItem> RpdFindAndReplaceItems { get; set; } = [];
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
        /// Последняя выбранная локация файлов УП
        /// </summary>
        [JsonInclude]
        public string CurriculumLastLocation { get; set; }
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
        /// Элементы для парсинга заголовков excel-файлов УП
        /// </summary>
        [JsonInclude]
        public List<CurriculumDisciplineHeader> CurriculumDisciplineParseItems { get; set; }
    }
}
