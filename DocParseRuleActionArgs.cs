using System.Text.RegularExpressions;
using Xceed.Document.NET;

namespace FosMan {
    /// <summary>
    /// Аргумент для кастомной функции Action правила парсинга
    /// </summary>
    public record DocParseRuleActionArgs<T> {
        /// <summary>
        /// Целевой объект
        /// </summary>
        public T Target { get; set; }
        /// <summary>
        /// Родительское правило
        /// </summary>
        public IDocParseRule<T> Rule { get; set; }
        /// <summary>
        /// Результат применения рег. выражения
        /// </summary>
        public Match Match { get; set; }
        /// <summary>
        /// Исходное значение из параграфа
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// Итоговое значение из документа
        /// </summary>
        public string Value { get; set; }
        /// <summary>
        /// Текущий параграф документа
        /// </summary>
        public Paragraph Paragraph { get; set; }
        /// <summary>
        /// Документ docx
        /// </summary>
        public Document Document { get; set; }
    }
}
