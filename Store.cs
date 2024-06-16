using System;
using System.Collections.Generic;
using System.Configuration.Internal;
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
        public const string FILE_NAME = "Store.json";

        Dictionary<string, Rpd> m_rpdDic = null;
        Dictionary<string, Fos> m_fosDic = null;

        /// <summary>
        /// Словарь РПД
        /// </summary>
        [JsonInclude]
        public Dictionary<string, Rpd> RpdDic { get => m_rpdDic; set => m_rpdDic = value; }

        /// <summary>
        /// Словарь ФОС
        /// </summary>
        [JsonInclude]
        public Dictionary<string, Fos> FosDic { get => m_fosDic; set => m_fosDic = value; }

        public Store() {
            m_rpdDic = [];
            m_fosDic = [];
        }
    }
}
