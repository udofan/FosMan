﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FosMan {
    public class CompetenceResult {
        //РОЗ УК-1.1:
        //- знать состав, структуру тре-буемых данных и информа-ции, процессы их сбора, обра-ботки и интерпретации; раз-личные варианты решения задачи.
        //Regex m_regexParseResult = new(@"(.+)[:\r\n]+(.+)", RegexOptions.Multiline | RegexOptions.Compiled);
        //static Regex m_regexParseResult = new(@"(.+\.\d{1,})([:\.\r\n ]|$)+(.*)", RegexOptions.Multiline | RegexOptions.Compiled);
        static Regex m_regexParseResult = new(@"(\S{3})[ -]+(.+\.\d{1,})([:\.\r\n ]|$)+(.*)", RegexOptions.Multiline | RegexOptions.Compiled);

        /// <summary>
        /// Код результата
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// Результат обучения
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Исходный текст результата
        /// </summary>
        public string SourceText { get; set; }
        
        /// <summary>
        /// Попытка парсинга результата
        /// </summary>
        /// <param name="text2"></param>
        static internal bool TryParse(string text, out CompetenceResult result) {
            result = new CompetenceResult() {
                SourceText = text
            };

            var match = m_regexParseResult.Match(text);

            if (match.Success) {
                //var code = string.Join("", match.Groups[1].Value.Trim(' ', ':', '\r', '\n').Split(' ')).ToUpper();

                result.Code = $"{match.Groups[1].Value} {match.Groups[2].Value}".ToUpper();

                //result.Code = $"{code[..3]} {result.Code}";
                result.Description = match.Groups[4].Value.Trim();
            }

            return match.Success;
        }

    }
}