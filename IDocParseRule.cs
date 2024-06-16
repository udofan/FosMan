using System.Text.RegularExpressions;

namespace FosMan {
    public interface IDocParseRule<T> {
        /// <summary>
        /// Флаг, что правило выключено
        /// </summary>
        //bool Disabled { get; set; }
        /// <summary>
        /// Тип парсинга: в текущей строке или захват до срабатывания маркера StopMarkers
        /// </summary>
        EParseType Type { get; set; }
        /// <summary>
        /// Строка, с которой будут конкатенироваться читаемые строки для Location = Multiline
        /// </summary>
        string MultilineConcatValue { get; set; }
        /// <summary>
        /// Имя свойства
        /// </summary>
        string PropertyName { get; set; }
        /// <summary>
        /// Тип свойства
        /// </summary>
        Type PropertyType { get; set; }
        /// <summary>
        /// Маркер начала захвата строк и индекс захвата при Location = Inline
        /// </summary>
        List<(Regex marker, int inlineGroupIdx)> StartMarkers { get; set; }
        /// <summary>
        /// Маркер останова сбора строк при Location = Multiline
        /// </summary>
        List<Regex> StopMarkers { get; set; }
        /// <summary>
        /// Список символов, по которым надо оттриммить полученную строку
        /// </summary>
        char[] TrimChars { get; set; }
        Action<T, Match, string> Action { get; set; }

        bool Equals<T>(IDocParseRule<T>? other);
        bool Equals(object? obj);
        int GetHashCode();
        string ToString();
    }
}