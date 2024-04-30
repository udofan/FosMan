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
        Regex m_regexParseIndicator = new(@"(.+\d{1})\.\s*([^\d{1}].+)$", RegexOptions.Compiled);

        //РОЗ УК-1.1:
        //- знать состав, структуру тре-буемых данных и информа-ции, процессы их сбора, обра-ботки и интерпретации; раз-личные варианты решения задачи.
        //Regex m_regexParseResult = new(@"(.+)[:\r\n]+(.+)", RegexOptions.Multiline | RegexOptions.Compiled);
        Regex m_regexParseResult2 = new(@"(.+\.\d{1,})([:\.\r\n ]|$)+(.*)", RegexOptions.Multiline | RegexOptions.Compiled);

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
        public string SourceIndicator { get; set; }
        /// <summary>
        /// Код результата
        /// </summary>
        public string ResultCode { get; set; }
        /// <summary>
        /// Результат обучения
        /// </summary>
        public string ResultDescription { get; set; }
        /// <summary>
        /// Исходный текст результата
        /// </summary>
        public string SourceResult { get; set; }

        /// <summary>
        /// Установка индикатора
        /// </summary>
        /// <param name="text1"></param>
        /// <exception cref="NotImplementedException"></exception>
        internal bool TryParseIndicator(string text) {
            SourceIndicator = text;

            var match = m_regexParseIndicator.Match(text);
            if (match.Success) {
                Code = string.Join("", match.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries)).ToUpper();
                Indicator = match.Groups[2].Value.Trim();
            }

            return match.Success;
        }

        /// <summary>
        /// Установка результата
        /// </summary>
        /// <param name="text2"></param>
        internal bool TryParseResult(string text) {
            SourceResult = text;

            var match = m_regexParseResult2.Match(text);

            if (match.Success) {
                var code = string.Join("", match.Groups[1].Value.Trim(' ', ':', '\r', '\n').Split(' ')).ToUpper();

                ResultCode = $"{code[..3]} {Code}";
                ResultDescription = match.Groups[3].Value.Trim();
            }
            
            return match.Success;
        }
    }
}
