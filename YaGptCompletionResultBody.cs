using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace FosMan {
    internal class YaGptCompletionResultBody {
        [JsonInclude]
        public List<YaGptCompletionAlternative> alternatives;
        [JsonInclude]
        public YaGptCompletionUsage usage;
        [JsonInclude]
        public string modelVersion;
    }
}
