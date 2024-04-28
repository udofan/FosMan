using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Достижение компетенции
    /// </summary>
    internal class CompetenceAchievement {
        /// <summary>
        /// Код
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// Индикатор
        /// </summary>
        public string Indicator { get; set; }
        /// <summary>
        /// Код результата
        /// </summary>
        public string ResultCode { get; set; }
        /// <summary>
        /// Результат обучения
        /// </summary>
        public string ResultDescription { get; set; }
    }
}
