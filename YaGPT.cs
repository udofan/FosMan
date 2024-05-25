using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace FosMan {
    internal class YaGPT {
        static string m_OAuthToken = "y0_AgAAAAAEx3K_AATuwQAAAAEFrx5AAADZrp-KaeNF2bZYCibNlcUk_7M_6Q";

        static readonly HttpClient m_httpClient = new HttpClient();

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
        }

        static async void RefreshIamToken() {
            bool refresh = false;
            try {
                var tokenAge = App.Config.YaGptIamTokenExpiresAt - DateTimeOffset.Now;

                if (string.IsNullOrEmpty(App.Config.YaGptIamToken) || tokenAge.TotalMinutes < 3) {
                    refresh = true;
                }

                //refresh = true;
                if (refresh) {
                    //App.Config.YaGptIamTokenRequestTimestamp = DateTime.Now;
                    //var values = new Dictionary<string, string> {
                    //    { "yandexPassportOauthToken", App.Config.YaOAuthToken }
                    //};

                    //var content = new FormUrlEncodedContent(values);

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
                    //    //"2024-05-25T19:29:33.514435084Z"
                        if (DateTimeOffset.TryParse(expiresValue.ToString(), out var dtOffset)) {
                            App.Config.YaGptIamTokenExpiresAt = dtOffset; //.DateTime;
                        }
                    }
 
                    App.SaveConfig();
                }
            }
            catch (Exception ex) {

            }
        }
    }
}
