using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FosMan {
    internal class CompetenceMatrixItem {
        //УК-1. Способен осуществлять поиск, критический анализ и синтез информации, применять системный подход для решения поставленных задач
        Regex m_parseText = new(@"(.*\d{1})\. (.*)", RegexOptions.Compiled);

        /// <summary>
        /// Код компетенции
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// Наименование
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Исходный текст
        /// </summary>
        public string SourceText { get; set; }

        /// <summary>
        /// Достижения компетенций
        /// </summary>
        public List<CompetenceAchievement> Achievements { get; set; }

        /// <summary>
        /// Инициализация базовых свойств по тексту
        /// </summary>
        /// <param name="text"></param>
        public bool InitFromText(string text) {
            SourceText = text;

            var match = m_parseText.Match(text);
            var result = match.Success;

            if (match.Success) {
                Code = string.Join("", match.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries));
                Title = match.Groups[2].Value.Trim();
            }

            return result;
        }
    }
}
