using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    internal class YandexGPT {
        string m_OAuthToken = "y0_AgAAAAAEx3K_AATuwQAAAAEFrx5AAADZrp-KaeNF2bZYCibNlcUk_7M_6Q";

        //curl -X POST \
        //-d '{"yandexPassportOauthToken":"<OAuth-токен>        "}' \
        //https://iam.api.cloud.yandex.net/iam/v1/tokens

        //export IAM_TOKEN =`yc iam create-token`
        //curl -X GET \
        //-H "Authorization: Bearer ${IAM_TOKEN}" \
        //https://resource-manager.api.cloud.yandex.net/resource-manager/v1/clouds
    }
}
