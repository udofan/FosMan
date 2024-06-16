using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static FosMan.Enums;

namespace FosMan {
    /// <summary>
    /// Элемент паспорта фонда оценочных средств текущего контроля, соотнесённых с индикаторами достижения компетенций
    /// </summary>
    internal class FosPassportItem {
        /// <summary>
        /// Тема
        /// </summary>
        public string Topic { get; set; }
        /// <summary>
        /// Список кодов контролируемого индикатора достижения компетенции
        /// </summary>
        public HashSet<string> CompetenceIndicators { get; set; }
        /// <summary>
        /// Наименования оценочного средства (сделано множественным на всяк)
        /// </summary>
        public HashSet<EEvaluationTool> EvaluationTools { get; set; }
    }
}
