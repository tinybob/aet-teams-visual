
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
            return response;
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