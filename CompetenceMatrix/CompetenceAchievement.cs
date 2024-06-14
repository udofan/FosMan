using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Достижение компетенции
    /// </summary>
    public class CompetenceAchievement {
        //УК-1.1. Определяет и ранжиру-ет информацию, требуемую для решения поставленных задач.
        static Regex m_regexParseIndicator = new(@"(.+\d{1})[\.]*\s*([^\d{1}].+)$", RegexOptions.Compiled);

        /// <summary>
        /// Код
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// Индикатор
        /// </summary>
        public string Indicator { get; set; }
        /// <summary>
        /// Исходный текст индикатора
        /// </summary>
        public string SourceText { get; set; }
        
        /// <summary>
        /// Список результатов обучения
        /// </summary>
        public List<CompetenceResult> Results { get; set; }

        /// <summary>
        /// Установка индикатора
        /// </summary>
        /// <param name="text1"></param>
        /// <exception cref="NotImplementedException"></exception>
        static internal bool TryParseIndicator(string text, out CompetenceAchievement achievement) {
            achievement = new() {
                SourceText = text, 
                Results = []
            };

            var match = m_regexParseIndicator.Match(text);
            if (match.Success) {
                achievement.Code = string.Join("", match.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries)).ToUpper();
                achievement.Indicator = match.Groups[2].Value.Trim();
            }

            return match.Success;
        }
    }
}
