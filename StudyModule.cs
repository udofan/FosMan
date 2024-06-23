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
        HashSet<string> m_competenceIndicators = null;

        /// <summary>
        /// Тема
        /// </summary>
        [JsonInclude]
        public string Topic { get; set; }
        /// <summary>
        /// Список кодов контролируемого индикатора достижения компетенции
        /// </summary>
        [JsonInclude]
        public HashSet<string> CompetenceIndicators {
            get => m_competenceIndicators;
            set => SetComptenceIndicators(value);
        }
        /// <summary>
        /// Наименования оценочного средства (сделано множественным на всяк)
        /// </summary>
        [JsonInclude]
        public HashSet<EEvaluationTool> EvaluationTools { get; set; }

        /// <summary>
        /// Установить список индикаторов с учетом нормализации значений
        /// </summary>
        /// <param name="items"></param>
        public void SetComptenceIndicators(IEnumerable<string> items) {
            if (items != null) {
                m_competenceIndicators = [];
                foreach (var item in items) {
                    if (CompetenceResult.TryParse(item, out var result)) {
                        m_competenceIndicators.Add(result.Code);
                    }
                }
            }
        }
    }
}
