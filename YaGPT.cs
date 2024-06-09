using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FosMan {
    internal class YaGpt {
        static string m_OAuthToken = "y0_AgAAAAAEx3K_AATuwQAAAAEFrx5AAADZrp-KaeNF2bZYCibNlcUk_7M_6Q";
        const string m_folderId = "b1gk7qgs2qql9gvkkn3r";

        static readonly HttpClient m_httpClient = new();

        //curl -X POST \
        //-d '{"yandexPassportOauthToken":"<OAuth-токен>        "}' \
        //https://iam.api.cloud.yandex.net/iam/v1/tokens

        //export IAM_TOKEN =`yc iam create-token`
        //curl -X GET \
        //-H "Authorization: Bearer ${IAM_TOKEN}" \
        //https://resource-manager.api.cloud.yandex.net/resource-manager/v1/clouds

        public static async void Init() {
            if (string.IsNullOrEmpty(App.Config.YaOAuthToken)) {
                App.Config.YaOAuthToken = m_OAuthToken;
            }

            await Task.Run(RefreshIamToken);

            //m_httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue().Add("Authorization", $"Bearer: {App.Config.YaGptIamToken}");
            m_httpClient.DefaultRequestHeaders.Add("x-folder-id", m_folderId);
        }

        /// <summary>
        /// Проверить и обновить IAM-токен
        /// </summary>
        static async Task<bool> RefreshIamToken() {
            bool refresh = false;
            try {
                var tokenAge = App.Config.YaGptIamTokenExpiresAt - DateTimeOffset.Now;

                if (string.IsNullOrEmpty(App.Config.YaGptIamToken) || tokenAge.TotalMinutes < 3) {
                    refresh = true;
                }

                if (refresh) {
                    using StringContent jsonContent = new(
                        JsonSerializer.Serialize(new {
                            yandexPassportOauthToken = App.Config.YaOAuthToken
                        }),
                        Encoding.UTF8,
                        "application/json");

                    var response = await m_httpClient.PostAsync("https://iam.api.cloud.yandex.net/iam/v1/tokens", jsonContent);
                    //var response = await Request()

                    var result = await response.Content.ReadAsStreamAsync();
                    
                    var json = JsonDocument.Parse(result);
                    if (json.RootElement.TryGetProperty("iamToken", out var tokenValue)) {
                        App.Config.YaGptIamToken = tokenValue.ToString();
                    }
                    if (json.RootElement.TryGetProperty("expiresAt", out var expiresValue)) {
                        //"2024-05-25T19:29:33.514435084Z"
                        if (DateTimeOffset.TryParse(expiresValue.ToString(), out var dtOffset)) {
                            App.Config.YaGptIamTokenExpiresAt = dtOffset;
                        }
                    }
 
                    App.SaveConfig();
                }
            }
            catch (Exception ex) {

            }
            return refresh;
        }

        public static async Task<YaGptCompletionResult> Completion(YaGptCompletionRequest request) {
            YaGptCompletionResult result = null;

            //RefreshIamToken();

            //request.modelUri = $"gpt://{m_folderId}/yandexgpt/latest";
            var json = App.SerializeObj(request);
            using StringContent jsonContent = new(json, Encoding.UTF8, "application/json");

            var requestMsg = new HttpRequestMessage() {
                Content = jsonContent,
                RequestUri = new("https://llm.api.cloud.yandex.net/foundationModels/v1/completion"),
                Method = HttpMethod.Post
            };
        
            var refreshResult = await RefreshIamToken();
            //if (refreshResult) {
                requestMsg.Headers.Authorization = new($"Bearer", App.Config.YaGptIamToken);
                //jsonContent.Headers.Add("Authorization", $"Bearer: {App.Config.YaGptIamToken}");
            //}
            //jsonContent.Headers.TryAddWithoutValidation("x-folder-id", m_folderId);

            var response = await m_httpClient.SendAsync(requestMsg); // //.PostAsync("https://llm.api.cloud.yandex.net/foundationModels/v1/completion", jsonContent);

            var responseResult = await response.Content.ReadAsStringAsync();
            result = App.Deserialize<YaGptCompletionResult>(responseResult);

            return result;
        }

        internal static async Task<string> TextGeneration(string systemText, string userText, double temp) {
            var resultText = "";

            var request = new YaGptCompletionRequest() {
                messages = new() {
                    new() {
                        role = "system",
                        text = systemText
                    },
                    new() {
                        role = "user",
                        text = userText
                    }
                },
                completionOptions = new() {
                    stream = false,
                    temperature = temp,
                    maxTokens = 8000
                },
                modelUri = $"gpt://{m_folderId}/yandexgpt/latest"
            };
            var result = await Completion(request);
            
            if (result?.result?.alternatives?.Any() ?? false) {
                resultText = result.result.alternatives.FirstOrDefault().message.text;
            }

            return resultText;
        }

        /// <summary>
        /// Получить список связанных дисциплин с указанной по списку возможных вариантов
        /// </summary>
        /// <param name="possibleDisciplineNames"></param>
        /// <param name="keyDisciplineName"></param>
        /// <returns></returns>
        internal static async Task<List<string>> GetRelatedDisciplines(string disciplineName, List<string> possibleDisciplines) {
            var resultList = new List<string>();
            var systemText = "Из заданного списка дисциплин определи только те, которые связаны с указанной дисциплиной. " +
                             "Ответ сформируй в формате JSON, где будет поле result со значением, где будут названия дисциплин.";
            var userText = $"Какие дисциплины связаны с дисциплиной \"{disciplineName}\" из следующего списка:\r\n{string.Join(",\r\n", possibleDisciplines)}";
            var result = await TextGeneration(systemText, userText, 0.0);
            if (!string.IsNullOrEmpty(result)) {
                //очистка ответа
                var posStart = result.IndexOf('{');
                var posLast = result.LastIndexOf('}');
                var json = result.Substring(posStart, posLast - posStart + 1);
                if (App.TryDeserialize(json, out YaGptDisciplineList list)) {
                    resultList = list.result;
                }
            }

            return resultList;
        }
    }
}
