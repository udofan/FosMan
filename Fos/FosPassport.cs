using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Описание паспорта ФОСа (3. Паспорт фонда оценочных средств текущего контроля, соотнесённых с индикаторами достижения компетенций)
    /// </summary>
    internal class FosPassport {
        public List<FosPassportItem> Items { get; set; }
    }
}
