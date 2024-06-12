using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

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

            await Task.Run(RefreshIamTokenAsync);

            //m_httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue().Add("Authorization", $"Bearer: {App.Config.YaGptIamToken}");
            m_httpClient.DefaultRequestHeaders.Add("x-folder-id", m_folderId);
        }

        /// <summary>
        /// Проверить и обновить IAM-токен
        /// </summary>
        static async Task<bool> RefreshIamTokenAsync() {
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
                var ttt = 0;
            }
            return refresh;
        }

        public static async Task<YaGptCompletionResult> CompletionAsync(YaGptCompletionRequest request) {
            YaGptCompletionResult result = null;
            try {

                //RefreshIamToken();

                //request.modelUri = $"gpt://{m_folderId}/yandexgpt/latest";
                var json = App.SerializeObj(request);
                using StringContent jsonContent = new(json, Encoding.UTF8, "application/json");

                var requestMsg = new HttpRequestMessage() {
                    Content = jsonContent,
                    RequestUri = new("https://llm.api.cloud.yandex.net/foundationModels/v1/completion"),
                    Method = HttpMethod.Post
                };

                var refreshResult = await RefreshIamTokenAsync();
                //if (refreshResult) {
                requestMsg.Headers.Authorization = new($"Bearer", App.Config.YaGptIamToken);
                //jsonContent.Headers.Add("Authorization", $"Bearer: {App.Config.YaGptIamToken}");
                //}
                //jsonContent.Headers.TryAddWithoutValidation("x-folder-id", m_folderId);

                var response = await m_httpClient.SendAsync(requestMsg); // //.PostAsync("https://llm.api.cloud.yandex.net/foundationModels/v1/completion", jsonContent);

                var responseResult = await response.Content.ReadAsStringAsync();
                //{"error":{"grpcCode":8,"httpCode":429,"message":"ai.textGenerationCompletionSessionsCount.count gauge quota limit exceed: allowed 1 requests","httpStatus":"Too Many Requests","details":[]}}
                if (!App.TryDeserialize<YaGptCompletionResult>(responseResult, out result)) {
                    //if (App.TryDeserialize<YaGptCompletionError>(responseResult, out var error)) {
                    var tt = 0;
                    //}
                }
            }
            catch (Exception ex) {
                var yy = 0;
            }
            return result;
        }

        ///// <summary>
        ///// Проверить и обновить IAM-токен
        ///// </summary>
        //static bool RefreshIamToken() {
        //    bool refresh = false;
        //    try {
        //        var tokenAge = App.Config.YaGptIamTokenExpiresAt - DateTimeOffset.Now;

        //        if (string.IsNullOrEmpty(App.Config.YaGptIamToken) || tokenAge.TotalMinutes < 3) {
        //            refresh = true;
        //        }

        //        if (refresh) {
        //            using StringContent jsonContent = new(
        //                JsonSerializer.Serialize(new {
        //                    yandexPassportOauthToken = App.Config.YaOAuthToken
        //                }),
        //                Encoding.UTF8,
        //                "application/json");

        //            var response = m_httpClient.PostAsync.Post("https://iam.api.cloud.yandex.net/iam/v1/tokens", jsonContent);
        //            //var response = await Request()

        //            var result = response.Content.ReadAsStream();

        //            var json = JsonDocument.Parse(result);
        //            if (json.RootElement.TryGetProperty("iamToken", out var tokenValue)) {
        //                App.Config.YaGptIamToken = tokenValue.ToString();
        //            }
        //            if (json.RootElement.TryGetProperty("expiresAt", out var expiresValue)) {
        //                //"2024-05-25T19:29:33.514435084Z"
        //                if (DateTimeOffset.TryParse(expiresValue.ToString(), out var dtOffset)) {
        //                    App.Config.YaGptIamTokenExpiresAt = dtOffset;
        //                }
        //            }

        //            App.SaveConfig();
        //        }
        //    }
        //    catch (Exception ex) {

        //    }
        //    return refresh;
        //}

        //public static YaGptCompletionResult Completion(YaGptCompletionRequest request) {
        //    YaGptCompletionResult result = null;

        //    //RefreshIamToken();

        //    //request.modelUri = $"gpt://{m_folderId}/yandexgpt/latest";
        //    var json = App.SerializeObj(request);
        //    using StringContent jsonContent = new(json, Encoding.UTF8, "application/json");

        //    var requestMsg = new HttpRequestMessage() {
        //        Content = jsonContent,
        //        RequestUri = new("https://llm.api.cloud.yandex.net/foundationModels/v1/completion"),
        //        Method = HttpMethod.Post
        //    };

        //    var refreshResult = await RefreshIamTokenAsync();
        //    //if (refreshResult) {
        //    requestMsg.Headers.Authorization = new($"Bearer", App.Config.YaGptIamToken);
        //    //jsonContent.Headers.Add("Authorization", $"Bearer: {App.Config.YaGptIamToken}");
        //    //}
        //    //jsonContent.Headers.TryAddWithoutValidation("x-folder-id", m_folderId);

        //    var response = await m_httpClient.SendAsync(requestMsg); // //.PostAsync("https://llm.api.cloud.yandex.net/foundationModels/v1/completion", jsonContent);

        //    var responseResult = await response.Content.ReadAsStringAsync();
        //    result = App.Deserialize<YaGptCompletionResult>(responseResult);

        //    return result;
        //}

        internal static async Task<(bool success, string text, YaGptCompletionErrorBody error)> TextGenerationAsync(string systemText, string userText, double temp) {
            (bool success, string text, YaGptCompletionErrorBody error) result = (false, null, null);
            
            try {

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
                var resultObj = await CompletionAsync(request);

                if (resultObj?.result?.alternatives?.Any() ?? false) {
                    result.success = true;
                    result.text = resultObj.result.alternatives.FirstOrDefault()?.message.text;
                }
                else if (resultObj?.error != null) {
                    result.error = resultObj?.error;
                }
            }
            catch (Exception ex) {
                var tt = 0;
            }
            return result;
        }

        /// <summary>
        /// Получить список связанных дисциплин с указанной по списку возможных вариантов
        /// </summary>
        /// <param name="possibleDisciplineNames"></param>
        /// <param name="keyDisciplineName"></param>
        /// <returns></returns>
        internal static async Task<List<string>> GetRelatedDisciplinesAsync(string disciplineName, List<string> possibleDisciplines) {
            var resultList = new List<string>();
            var systemText = "Из заданного списка дисциплин определи только те, которые связаны с указанной дисциплиной. " +
                             "Ответ сформируй в формате JSON, где будет поле result со значением, где будут названия дисциплин.";
            var userText = $"Какие дисциплины связаны с дисциплиной \"{disciplineName}\" из следующего списка:\r\n{string.Join(",\r\n", possibleDisciplines)}";
            //Debug.WriteLine($"TextGenerationAsync call...");
            var result = await TextGenerationAsync(systemText, userText, 0.0);
            //Debug.WriteLine($"TextGenerationAsync done.");
            if (result.success && !string.IsNullOrEmpty(result.text)) {
                var text = result.text;
                //очистка ответа
                var posStart = text.IndexOf('{');
                var posLast = text.LastIndexOf('}');
                var json = text.Substring(posStart, posLast - posStart + 1);
                if (App.TryDeserialize(json, out YaGptDisciplineList list)) {
                    resultList = list.result;
                }
            }
            else {
                Thread.Sleep(1000);
                resultList = await GetRelatedDisciplinesAsync(disciplineName, possibleDisciplines);
                //throw new Exception(result.error.message);
            }

            return resultList;
        }

        /// <summary>
        /// Получить список связанных дисциплин с указанной по списку возможных вариантов
        /// </summary>
        /// <param name="possibleDisciplineNames"></param>
        /// <param name="keyDisciplineName"></param>
        /// <returns></returns>
        internal static List<string> GetRelatedDisciplines(string disciplineName, List<string> possibleDisciplines) {
            var resultList = new List<string>();
            var task = GetRelatedDisciplinesAsync(disciplineName, possibleDisciplines);
            //task.Wait();
            resultList = task.GetAwaiter().GetResult();
            //resultList = task.Result;

            return resultList;
        }
    }
}
