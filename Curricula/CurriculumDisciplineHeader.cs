using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace FosMan {
    public enum EPropertyTestFunction {
        Contains,
        Equals,
        StartsWith
    }

    /// <summary>
    /// Описание заголовка таблицы дисциплин (лист План)
    /// </summary>
    internal class CurriculumDisciplineHeader {
        Type m_type = null;
        string m_typeName = null;

        /// <summary>
        /// Тестовый текст для выявления колонки свойства
        /// </summary>
        [JsonInclude]
        public string Text { get; set; }
        /// <summary>
        /// Функция для тестирования значения колонки (исп. вместе с HeaderText)
        /// </summary>
        [JsonInclude]
        public EPropertyTestFunction TestFunction { get; set; }
        /// <summary>
        /// Имя целевого свойства (класса CurriculumDiscipline)
        /// </summary>
        [JsonInclude]
        public string TargetProperty { get; set; }
        /// <summary>
        /// Целевой тип
        /// </summary>
        [JsonIgnore]
        public Type TargetType {
            get {
                if (m_type == null) {
                    m_type = Type.GetType(m_typeName);
                }
                return m_type;
            }
            set { m_type = value; }
        }
        /// <summary>
        /// Имя целевого типа
        /// </summary>
        [JsonInclude]
        public string TargetTypeName { 
            get => m_typeName; 
            set => m_typeName = value;
        }
        /// <summary>
        /// Номер совпадения (для одинаковых заголовков)
        /// </summary>
        [JsonInclude]
        public int MatchNumber { get; set; } = 0;
        /// <summary>
        /// Свойство элемента списка (исп. с парой PropertyIndex)
        /// </summary>
        [JsonInclude]
        public string SubProperty { get; set; } = null;
        /// <summary>
        /// Индекс для свойства типа Array
        /// </summary>
        [JsonInclude]
        public int PropertyIndex { get; set; } = -1;

        /// <summary>
        /// Проверка: является ли текст подходящим для заголовка
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public bool Match(string text) {
            var match = false;

            text = text.Trim();

            if (TestFunction == EPropertyTestFunction.Contains) {
                match = text.Contains(Text, StringComparison.CurrentCultureIgnoreCase);
            }
            else if (TestFunction == EPropertyTestFunction.Equals) {
                match = text.Equals(Text, StringComparison.CurrentCultureIgnoreCase);
            }
            else if (TestFunction == EPropertyTestFunction.StartsWith) {
                match = text.StartsWith(Text, StringComparison.CurrentCultureIgnoreCase);
            }

            return match;
        }
    }
}
