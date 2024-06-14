using System.Security.Policy;
using System.Text.Json.Serialization;

namespace FosMan
{
    internal class YaGptCompletionUsage
    {
        [JsonInclude]
        public string inputTextTokens;
        [JsonInclude]
        public string completionTokens;
        [JsonInclude]
        public string totalTokens;
    }
}