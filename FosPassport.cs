using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    internal class FosPassport {
        /// <summary>
        /// Тема
        /// </summary>
        public string Topic { get; set; }
        /// <summary>
        /// Код контролируемого индикатора достижения компетенции
        /// </summary>
        public string CompetenceIndicator { get; set; }
        public EEvaluationTool EvaluationTool { get; set; }
    }
}
