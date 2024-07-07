using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FosMan {
    public class CompetenceMatrixItem : BaseObj {
        //УК-1. Способен осуществлять поиск, критический анализ и синтез информации, применять системный подход для решения поставленных задач
        //УК-1
        //Способен осуществлять поиск, критический анализ и синтез информации, применять системный подход для решения поставленных задач.
        static Regex m_parseCode = new(@"([а-яА-Я]{2,}.*\d+)", RegexOptions.Compiled);
        static Regex m_parseText = new(@"([а-яА-Я]{2,}.*\d+)([\.\r\n\s{1}$]|$)(.*)$", RegexOptions.Compiled);

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
        /// Этап формирования компетенций (семестр) [исп. в ФОС п. 2.1]
        /// </summary>
        public int Semester { get; set; } = -1;

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
                matrixItem.Code = NormalizeCode(match.Groups[1].Value);
                matrixItem.Title = match.Groups[3].Value.Trim();
            }

            return result;
        }

        /// <summary>
        /// Проверка, что переданный текст содержит код компетенции
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static bool TestCode(string text) => m_parseCode.IsMatch(text);

        /// <summary>
        /// Нормализация кода компетенции
        /// </summary>
        /// <param name="code"></param>
        public static string NormalizeCode(string code) => string.Join("", code.Split(' ', StringSplitOptions.RemoveEmptyEntries)).ToUpper().Trim(' ', '.');

        /// <summary>
        /// Поиск достижения по коду
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        public CompetenceAchievement FindAchievement(string code) {
            var normalizedCode = CompetenceMatrixItem.NormalizeCode(code);

            return Achievements.FirstOrDefault(a => a.Code.Equals(normalizedCode));
        }
    }
}
