// https://medium.com/@fiqriismail/how-to-get-an-access-token-for-microsoft-graph-api-using-node-js-258723f29cc6
const APP_ID = "b5615dbe-0af5-49fd-ab09-803e91be7bd9";
const APP_SECERET = "L9c1qlg8x1CfH8StSyfVtkB23vD-C~-.x.";
const TOKEN_ENDPOINT =
  "https://login.microsoftonline.com/b4b87e8f-df64-41ff-9ba4-a4930ebc804b/oauth2/v2.0/token";
const MS_GRAPH_SCOPE = "https://graph.microsoft.com/.default";

const axios = require('axios');
const qs = require('qs');

const postData = {
  client_id: APP_ID,
  scope: MS_GRAPH_SCOPE,
  client_secret: APP_SECERET,
  grant_type: 'client_credentials'
};

axios.defaults.headers.post['Content-Type'] =
  'application/x-www-form-urlencoded';

let token = '';

axios
  .post(TOKEN_ENDPOINT, qs.stringify(postData))
  .then(response => {
    console.log(response.data.access_token);
  })
  .catch(error => {
    console.log(error);
  });