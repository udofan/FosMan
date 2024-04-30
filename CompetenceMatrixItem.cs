using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FosMan {
    public class CompetenceMatrixItem {
        //УК-1. Способен осуществлять поиск, критический анализ и синтез информации, применять системный подход для решения поставленных задач
        //УК-1
        //Способен осуществлять поиск, критический анализ и синтез информации, применять системный подход для решения поставленных задач.
        static Regex m_parseText = new(@"(.*\d+)([\.\r\n\s{1}$]|$)(.*)$", RegexOptions.Compiled);

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
        /// Попытка отпарсить текст с кодом компетенции и описанием
        /// </summary>
        /// <param name="text"></param>
        public static bool TryParse(string text, out CompetenceMatrixItem matrixItem) {
            matrixItem = new CompetenceMatrixItem() {
                SourceText = text, 
                Achievements = []
            };

            var match = m_parseText.Match(text);
            var result = match.Success;

            if (match.Success) {
                matrixItem.Code = string.Join("", match.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries)).ToUpper();
                matrixItem.Title = match.Groups[3].Value.Trim();
            }

            return result;
        }
    }
}
