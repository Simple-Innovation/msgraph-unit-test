import { request } from 'http';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"

import 'isomorphic-fetch';

test('upload small file', () => {
    const accessToken = "DDD";

    let url = "https://graph.microsoft.com/v1.0/me/messages";
    let request = new Request(url, {
        method: "GET",
        headers: new Headers({
            "Authorization": "Bearer " + accessToken
        })
    });
    
    fetch(request)
    .then((response) => {
        response.json().then((res) => {
            let messages:[MicrosoftGraph.Message] = res.value;
            for (let msg of messages) { //iterate through the recent messages
                console.log(msg.subject);
                console.log(msg.toRecipients[0].emailAddress.address);
            }
        });
    
    })
    .catch((error) => {
        console.error(error);
    });
})
