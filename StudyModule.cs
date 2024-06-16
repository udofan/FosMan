using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using static FosMan.Enums;

namespace FosMan {
    /// <summary>
    /// Описание модуля обучения: раздел (тема), оценочное средство, компетенции, время?
    /// </summary>
    internal class StudyModule {
        /// <summary>
        /// Тема
        /// </summary>
        [JsonInclude]
        public string Topic { get; set; }
        /// <summary>
        /// Список кодов контролируемого индикатора достижения компетенции
        /// </summary>
        [JsonInclude]
        public HashSet<string> CompetenceIndicators { get; set; }
        /// <summary>
        /// Наименования оценочного средства (сделано множественным на всяк)
        /// </summary>
        [JsonInclude]
        public HashSet<EEvaluationTool> EvaluationTools { get; set; }
    }
}
