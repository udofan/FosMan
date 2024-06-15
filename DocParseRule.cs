using System.Text.RegularExpressions;

namespace FosMan {
    /// <summary>
    /// Описание правила парсинга значения свойства объекта
    /// </summary>
    //public record DocParseRule : IDocParseRule {
    //    /// <summary>
    //    /// Имя свойства
    //    /// </summary>
    //    public string Name { get; set; }
    //    /// <summary>
    //    /// Маркер начала захвата строк и индекс захвата при Location = Inline
    //    /// </summary>
    //    public List<(Regex marker, int inlineGroupIdx)> StartMarkers { get; set; }
    //    /// <summary>
    //    /// Номер группы в маркере начала для Location = Inline
    //    /// </summary>
    //    //public int InlineGroupIndex { get; set; }
    //    /// <summary>
    //    /// Тип парсинга: в текущей строке или захват до срабатывания маркера StopMarkers
    //    /// </summary>
    //    public EParseType Type { get; set; } = EParseType.Inline;
    //    /// <summary>
    //    /// Маркер останова сбора строк при Location = Multiline
    //    /// </summary>
    //    public List<Regex> StopMarkers { get; set; } = null;
    //    /// <summary>
    //    /// Строка, с которой будут конкатенироваться читаемые строки для Location = Multiline
    //    /// </summary>
    //    public string MultilineConcatValue { get; set; } = " ";
    //    /// <summary>
    //    /// Список символов, по которым надо оттриммить полученную строку
    //    /// </summary>
    //    public char[] TrimChars { get; set; } = null;
    //}
}
