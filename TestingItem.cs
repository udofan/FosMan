using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace FosMan {
    public class TestingItem : BaseObj {
        /// <summary>
        /// Вопрос
        /// </summary>
        [JsonInclude]
        public string Question { get; set; }
        /// <summary>
        /// Список ответов
        /// </summary>
        [JsonInclude]
        public List<string> Answers { get; set; }
        /// <summary>
        /// Индексы правильных ответов
        /// </summary>
        [JsonInclude]
        public List<int> RightAnswerIndicies { get; set; }
    }
}
