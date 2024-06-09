using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// База данных
    /// </summary>
    internal class Store {
        Dictionary<string, Rpd> m_rpdDic = null;

        /// <summary>
        /// Словарь РПД
        /// </summary>
        [JsonInclude]
        public Dictionary<string, Rpd> RpdDic { get => m_rpdDic; set => m_rpdDic = value; }

        public Store() {
            m_rpdDic = [];
        }
    }
}
