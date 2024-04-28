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

    internal class CurriculumProperty {
        /// <summary>
        /// Тестовый текст для выявления колонки свойства
        /// </summary>
        public string HeaderText { get; set; }
        /// <summary>
        /// Функция для тестирования значения колонки (исп. вместе с HeaderText)
        /// </summary>
        public EPropertyTestFunction TestFunction { get; set; }
        /// <summary>
        /// Имя целевого свойства (класса CurriculumDiscipline)
        /// </summary>
        public string TargetProperty { get; set; }
    }
}
