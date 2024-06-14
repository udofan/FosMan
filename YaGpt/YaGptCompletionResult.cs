using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace FosMan
{
    internal class YaGptCompletionResult
    {
        [JsonInclude]
        public YaGptCompletionResultBody result { get; set; }
        [JsonInclude]
        public YaGptCompletionErrorBody error { get; set; }
    }
}
