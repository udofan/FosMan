using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    public enum EPropertyTestFunction {
        Contains,
        Equals
    }

    internal class CurriculumDisciplineHeader {
        /// <summary>
        /// Тестовый текст для выявления колонки свойства
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// Функция для тестирования значения колонки (исп. вместе с HeaderText)
        /// </summary>
        public EPropertyTestFunction TestFunction { get; set; }
        /// <summary>
        /// Имя целевого свойства (класса CurriculumDiscipline)
        /// </summary>
        public string TargetProperty { get; set; }
        /// <summary>
        /// Целевой тип
        /// </summary>
        public Type TargetType { get; set; }
        /// <summary>
        /// Номер совпадения (для одинаковых заголовков)
        /// </summary>
        public int MatchNumber { get; set; } = 0;

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

            return match;
        }
    }
}
