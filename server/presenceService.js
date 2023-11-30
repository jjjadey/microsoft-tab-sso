import fetch from 'node-fetch';

export async function getPresence(accessToken) {

    const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me/presence",
        {
            method: 'GET',
            headers: {
                "accept": "application/json",
                "authorization": `bearer ${accessToken}`
            },
            mode: 'cors',
            cache: 'default'
        });
    if (!graphResponse.ok) {
        throw (`Error ${graphResponse.status} calling Microsoft Graph: ${graphResponse.statusText}`);
    }
    const profile = await graphResponse.json();
    return profile;
}