using System.Text.Json.Serialization;

namespace FosMan
{
    internal class YaGptCompletionAlternative
    {
        [JsonInclude]
        public YaGptCompletionMessage message;
        [JsonInclude]
        public string status;
    }
}