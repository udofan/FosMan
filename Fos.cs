using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Описание ФОСа
    /// </summary>
    internal class Fos {
        /// <summary>
        /// Дисциплина
        /// </summary>
        public string Discipline { get; set; }
        /// <summary>
        /// Форма обучения
        /// </summary>
        public string FormOfEducation { get; set; }
        /// <summary>
        /// Дата утверждения
        /// </summary>
        public DateTime ApprovalDate { get; set; }
        /// <summary>
        /// Направление подготовки
        /// </summary>
        public string DirectionOfPreparation { get; set; }
        /// <summary>
        /// Профиль
        /// </summary>
        public string Profile { get; set; }
        /// <summary>
        /// Компетенции (коды)
        /// </summary>
        public List<string> Competencies { get; set; }
        /// <summary>
        /// 3. Паспорт фонда оценочных средств текущего контроля, соотнесённых с индикаторами достижения компетенций
        /// </summary>
        public List<FosPassport> Passport { get; set; }
    }
}
