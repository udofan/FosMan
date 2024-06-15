using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static FosMan.Enums;

namespace FosMan {
    /// <summary>
    /// Паспорт фонда оценочных средств текущего контроля, соотнесённых с индикаторами достижения компетенций
    /// </summary>
    internal class FosPassport {
        /// <summary>
        /// Тема
        /// </summary>
        public string Topic { get; set; }
        /// <summary>
        /// Список кодов контролируемого индикатора достижения компетенции
        /// </summary>
        public HashSet<string> CompetenceIndicators { get; set; }
        /// <summary>
        /// Наименование оценочного средства
        /// </summary>
        public EEvaluationTool EvaluationTool { get; set; }
    }
}
