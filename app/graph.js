/** 
 * Helper function to call MS Graph API endpoint
 * using the authorization bearer token scheme
*/
function callMSGraph(endpoint, token, callback, isFetchingAll) {
    const headers = new Headers();
    const bearer = `Bearer ${token}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    console.log('request made to Graph API at: ' + new Date().toString());

    fetch(endpoint, options)
        .then(response => response.json())
        .then(async response => {
            if(response['@odata.nextLink'] && isFetchingAll) {
                const content = await fetchNext(response['@odata.nextLink'], options);
                response.value = response.value.concat(content.value);
            }
            cacheData(endpoint, response);
            return response;
        })
        .then(response => callback(response, endpoint))
        .catch(error => console.log(error));
}

function postMSGraph(endpoint, token, body, callback) {
    const headers = new Headers();
    const bearer = `Bearer ${token}`;

    headers.append("Authorization", bearer);
    headers.append("Content-Type", "application/json");

    const options = {
        method: "POST",
        headers: headers,
        body: body
    };

    console.log('request made to Graph API at: ' + new Date().toString());

    fetch(endpoint, options)
        .then(response => {
            if(response.ok && response.statusText == 'No Content')
                return 'ok';
            else 
                return response.json();
        })
        .then(response => callback(response, endpoint))
        .catch(error => console.log(error));
}

function patchMSGraph(endpoint, token, body, callback) {
    const headers = new Headers();
    const bearer = `Bearer ${token}`;

    headers.append("Authorization", bearer);
    headers.append("Content-Type", "application/json");

    const options = {
        method: "PATCH",
        headers: headers,
        body: body
    };

    console.log('request made to Graph API at: ' + new Date().toString());

    fetch(endpoint, options)
        .then(response => {
            if(response.ok && response.statusText == 'No Content')
                return 'ok';
            else 
                return response.json();
        })
        .then(response => callback(response, endpoint))
        .catch(error => console.log(error));
}

async function fetchNext(endpoint, options) {
    return fetch(endpoint, options)
        .then(response => response.json())
        .then(async response => {
            if(response['@odata.nextLink']) {
                const next = await fetchNext(response['@odata.nextLink'], options);
                response.value = response.value.concat(next.value);
            }

            return response;
        })
}

function cacheData(endpoint, data) {
    // let events = sessionStorage.getItem('teams_event');
    // let messages = sessionStorage.getItem('teams_message');
    if((endpoint.indexOf('events') < 0 && endpoint.indexOf('messages') < 0)
        || sessionStorage.getItem(endpoint) || data.value.length == 0)
        return;

    const value = JSON.stringify(data.value);
    sessionStorage.setItem(endpoint, value);
}