const axios = require("axios");
const qs = require("qs");

const APP_ID = "b5615dbe-0af5-49fd-ab09-803e91be7bd9";
const APP_SECERET = "L9c1qlg8x1CfH8StSyfVtkB23vD-C~-.x.";
const TOKEN_ENDPOINT =
  "https://login.microsoftonline.com/b4b87e8f-df64-41ff-9ba4-a4930ebc804b/oauth2/v2.0/token";
const MS_GRAPH_SCOPE = "https://graph.microsoft.com/.default";

class ClientCredentialAuthenticationProvider {
  async getAccessToken() {
    const url = TOKEN_ENDPOINT;

    const body = {
      client_id: APP_ID,
      client_secret: APP_SECERET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    };
    axios.defaults.headers.post["Content-Type"] =
      "application/x-www-form-urlencoded";
    try {
      let response = await axios.post(url, qs.stringify(body))

      if (response.status == 200) {

          return response.data.access_token;

      } else {

          throw new Error("Non 200OK response on obtaining token...")

      }
    }
    catch (error) {

        throw new Error("Error on obtaining token...")

    }
  }
}
module.exports = ClientCredentialAuthenticationProvider;
