// https://www.npmjs.com/package/@microsoft/microsoft-graph-client
const mqtt = require("mqtt");
const ClientCredentialAuthenticationProvider = require("./AuthenticationProvider");
var graph = require("@microsoft/microsoft-graph-client");
const common = require('@bgroves/common');
require("isomorphic-fetch");
var mqttClient;

var { MQTT_SERVER,MQTT_PORT } = process.env;

async function main() 
{
    const clientOptions = {
        defaultVersion: "v1.0",
        debugLogging: false,
        authProvider: new ClientCredentialAuthenticationProvider()
    }

    const client = graph.Client.initWithMiddleware(clientOptions);

const mail =  {
    message: {
      subject: "Tooling Issue",
      body: {
		content: "<h5>Tooling Issue</h5><br>Tool 1 on CNC 120 which is running Dana RH 6K Knuckles is not reaching tool life.<br>If you are at work you can go to the <a href='https://eng'>app</a>",
		contentType: "html",
        // contentType: "Text",
        // content: "Tool 1 on Dana RH 6K Knuckles is not reaching tool life."
      },
      toRecipients: [
        {
          emailAddress: {
            address: "bgroves@mobexglobal.com"
          }
        }
      ],
    //   ccRecipients: [
    //     {
    //       emailAddress: {
    //         address: "dkreps@mobexglobal.com"
    //       }
    //     }
    //   ]
    },
    saveToSentItems: "false"
  };      
//   https://docs.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http
// https://docs.microsoft.com/en-us/graph/tutorials/node?tutorial-step=3
// https://www.npmjs.com/package/@microsoft/microsoft-graph-client
// https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/HEAD/docs/OtherAPIs.md
// let res = await client.api('/users/bgroves@buschegroup.com/sendMail')
// .post(mail);

try {
  common.log(`Starting Alarms13319`);
  common.log(`MQTT_SERVER=${MQTT_SERVER},MQTT_PORT=${MQTT_PORT}`);
  const mqttClient = mqtt.connect(`mqtt://${MQTT_SERVER}:${MQTT_PORT}`);

  mqttClient.on("connect", function () {
    mqttClient.subscribe("presence", function (err) {
      if (!err) {
        mqttClient.publish("presence", "Hello mqtt");
      }
    });
  });
} catch (err) {
  console.log("Error !!!", err);
}


}
main();
/*
https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/HEAD/docs/QueryParameters.md
https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/HEAD/docs/Actions.md
const request = await client
.api("/users")
// https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/HEAD/docs/QueryParameters.md
.select("id, displayName, mail,userPrincipalName,ID")
// https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/HEAD/docs/OtherAPIs.md
//		.query("$select=displayName")
.get()
.catch((error) => {
  console.log(error);
});
console.log(request.value);
*/

/*
Use Alarm13319
Check for issues matching each subscription record.
Send mail if found.
*/