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
    public class StudyModule {
        HashSet<string> m_competenceResults = null;

        /// <summary>
        /// Тема
        /// </summary>
        [JsonInclude]
        public string Topic { get; set; }
        /// <summary>
        /// Список кодов результатов достижения компетенций
        /// </summary>
        [JsonInclude]
        public HashSet<string> CompetenceResultCodes {
            get => m_competenceResults;
            set => SetCompetenceResults(value);
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
        public void SetCompetenceResults(IEnumerable<string> items) {
            if (items != null) {
                m_competenceResults = [];
                foreach (var item in items) {
                    if (CompetenceResult.TryParse(item, out var result)) {
                        m_competenceResults.Add(result.Code);
                    }
                }
            }
        }
    }
}
