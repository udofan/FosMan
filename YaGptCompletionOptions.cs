using System.Text.Json.Serialization;

namespace FosMan {
    public class YaGptCompletionOptions {
        [JsonInclude]
        public bool stream;
        [JsonInclude]
        public double temperature;
        [JsonInclude]
        public int maxTokens;
    }
}