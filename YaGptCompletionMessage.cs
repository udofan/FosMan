using System.Text.Json.Serialization;

namespace FosMan {
    public class YaGptCompletionMessage {
        [JsonInclude]
        public string role;
        [JsonInclude]
        public string text;
    }
}