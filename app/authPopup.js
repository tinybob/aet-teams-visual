// Create the main myMSALObj instance
// import * as msal from '../node_modules/@azure/msal-browser';

// configuration parameters are located at authConfig.js
const myMSALObj = new msal.PublicClientApplication(msalConfig);

let username = "";

function selectAccount() {

    /**
     * See here for more info on account retrieval: 
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */

    const currentAccounts = myMSALObj.getAllAccounts();
    if (currentAccounts.length === 0) {
        return;
    } else if (currentAccounts.length > 1) {
        // Add choose account code here
        console.warn("Multiple accounts detected.");
    } else if (currentAccounts.length === 1) {
        username = currentAccounts[0].username;
        showWelcomeMessage(username);
    }
}

function handleResponse(response) {

    /**
     * To see the full list of response object properties, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#response
     */

    if (response !== null) {
        username = response.account.username;
        showWelcomeMessage(username);


    } else {
        selectAccount();
    }
}

function signIn() {

    /**
     * You can pass a custom request object below. This will override the initial configuration. For more information, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#request
     */

    myMSALObj.loginPopup()
        .then(handleResponse)
        .catch(error => {
            console.error(error);
        });
}

function signOut() {

    /**
     * You can pass a custom request object below. This will override the initial configuration. For more information, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#request
     */

    // Choose which account to logout from by passing a username.

    const logoutRequest = {
        account: myMSALObj.getAccountByUsername(username)
    };

    myMSALObj.logout(logoutRequest);
}

function getTokenPopup(request) {

    /**
     * See here for more info on account retrieval: 
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */
    request.account = myMSALObj.getAccountByUsername(username);
    
    return myMSALObj.acquireTokenSilent(request)
        .catch(error => {
            console.warn("silent token acquisition fails. acquiring token using popup");
            if (error instanceof msal.InteractionRequiredAuthError) {
                // fallback to interaction when silent call fails
                return myMSALObj.acquireTokenPopup(request)
                    .then(tokenResponse => {
                        console.log(tokenResponse);
                        return tokenResponse;
                    }).catch(error => {
                        console.error(error);
                    });
            } else {
                console.warn(error);   
            }
    });
}

function seeProfile() {
    if(localStorage.length) {
        for (let index = 0; index < localStorage.length; index ++) {
            key = localStorage.key(index);
            value = localStorage.getItem(key);
            console.log(`key: ${key}, value: ${value}`)
        }
    }

    if(sessionStorage.length) {
        for (let index = 0; index < sessionStorage.length; index ++) {
            key = sessionStorage.key(index);
            value = sessionStorage.getItem(key);
            console.log(`key: ${key}, value: ${value}`)
        }
    }
    getTokenPopup(loginRequest)
        .then(response => {
            callMSGraph(graphConfig.graphMeEndpoint, response.accessToken, updateUI);
        }).catch(error => {
            console.error(error);
        });
}

function readMail() {
    getTokenPopup(tokenRequest)
        .then(response => {
            callMSGraph(graphConfig.graphMailEndpoint, response.accessToken, updateUI);
        }).catch(error => {
            console.error(error);
        });
}

function getPerson() {
    getTokenPopup(tokenRequest)
        .then(response => {
            // const endpoint = graphConfig.graphEndpoint + `/users/43e97f91-9324-43b1-b104-d9c7e4a6784d/people`
            const endpoint = graphConfig.graphEndpoint + '/me/people';
            callMSGraph(endpoint, response.accessToken, updateUI)
        }).catch(error => {
            console.error(error);
        })
}

function getTeams() {
    getTokenPopup(teamsRequest)
        .then(response => {
            let url = graphConfig.graphMeEndpoint + `/joinedTeams`;
            callMSGraph(url, response.accessToken, updateUI, true);
            // callMSGraph(url, response.accessToken, (data, endpoint) => {

            //     data.value.map(group => {
            //         const channelEndpoint = graphConfig.graphBetaEndpoint + `/teams/${group.id}/channels`;
                    
            //         // Group events
            //         const eventsEndpoint = graphConfig.graphBetaEndpoint + `/groups/${group.id}/events`;
            //         callMSGraph(eventsEndpoint, response.accessToken, (events, endpoint) => {
            //             group.events = events;
            //         })
            //         callMSGraph(channelEndpoint, response.accessToken, (channels, endpoint) => {
            //             group.channels = channels.value;
            //             group.channels.map( channel => {
            //                 const messageEndpoint = channelEndpoint + `/${channel.id}/messages`;
            //                 // response.accessToken = 'eyJ0eXAiOiJKV1QiLCJub25jZSI6Il9LVDhzeHNhQVFLWHd3VEFpRXEyaVlQU0dJQzA3bFpGYjRNYTE4b2otWmciLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC85M2YzMzU3MS01NTBmLTQzY2YtYjA5Zi1jZDMzMTMzOGQwODYvIiwiaWF0IjoxNjE4MzA3MTUyLCJuYmYiOjE2MTgzMDcxNTIsImV4cCI6MTYxODMxMTA1MiwiYWNjdCI6MCwiYWNyIjoiMSIsImFjcnMiOlsidXJuOnVzZXI6cmVnaXN0ZXJzZWN1cml0eWluZm8iLCJ1cm46bWljcm9zb2Z0OnJlcTEiLCJ1cm46bWljcm9zb2Z0OnJlcTIiLCJ1cm46bWljcm9zb2Z0OnJlcTMiLCJjMSIsImMyIiwiYzMiLCJjNCIsImM1IiwiYzYiLCJjNyIsImM4IiwiYzkiLCJjMTAiLCJjMTEiLCJjMTIiLCJjMTMiLCJjMTQiLCJjMTUiLCJjMTYiLCJjMTciLCJjMTgiLCJjMTkiLCJjMjAiLCJjMjEiLCJjMjIiLCJjMjMiLCJjMjQiLCJjMjUiXSwiYWlvIjoiRTJaZ1lIanozaVpYL0F1SGZ3bnZoK3YvdXFzT0xWQmtkTkRjMVppOXpOTWs2Yy9sazNvQSIsImFtciI6WyJwd2QiXSwiYXBwX2Rpc3BsYXluYW1lIjoiR3JhcGggZXhwbG9yZXIgKG9mZmljaWFsIHNpdGUpIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IlN1biIsImdpdmVuX25hbWUiOiJZdWJvIiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMjcuMTguODEuNSIsIm5hbWUiOiJTdW4sIFl1Ym8gKEJvYiwgRVMtQXBwcy1HRC1DSElOQS1XSCkiLCJvaWQiOiJlZDMwMDBhNi0yZjQzLTRiMDUtYWViYS01YTQ1YWY4MWY5OWYiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMjcxODcxMjg5My00MjU3ODkzMTAwLTM3ODY4MzU3Ni0yNDE1MTQiLCJwbGF0ZiI6IjUiLCJwdWlkIjoiMTAwMzdGRkVBNzg4OUU4MyIsInJoIjoiMC5BUTBBY1RYemt3OVZ6ME93bjgwekV6alFoclhJaTk3NTJiRklxSzIzU05weVVHUU5BTE0uIiwic2NwIjoiQWNjZXNzUmV2aWV3LlJlYWRXcml0ZS5BbGwgQXVkaXRMb2cuUmVhZC5BbGwgQ2FsZW5kYXJzLlJlYWQgQ2FsZW5kYXJzLlJlYWRXcml0ZSBDb250YWN0cy5SZWFkV3JpdGUgRGV2aWNlTWFuYWdlbWVudEFwcHMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudENvbmZpZ3VyYXRpb24uUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlJlYWQuQWxsIERldmljZU1hbmFnZW1lbnRTZXJ2aWNlQ29uZmlnLlJlYWQuQWxsIERpcmVjdG9yeS5BY2Nlc3NBc1VzZXIuQWxsIERpcmVjdG9yeS5SZWFkV3JpdGUuQWxsIGVtYWlsIEV4dGVybmFsSXRlbS5SZWFkLkFsbCBGaWxlcy5SZWFkV3JpdGUuQWxsIEdyb3VwLlJlYWRXcml0ZS5BbGwgSWRlbnRpdHlSaXNrRXZlbnQuUmVhZC5BbGwgTWFpbC5SZWFkLlNoYXJlZCBNYWlsLlJlYWRXcml0ZSBNYWlsLlJlYWRXcml0ZS5TaGFyZWQgTWFpbGJveFNldHRpbmdzLlJlYWQgTm90ZXMuUmVhZFdyaXRlLkFsbCBvcGVuaWQgUGVvcGxlLlJlYWQgcHJvZmlsZSBSZXBvcnRzLlJlYWQuQWxsIFNpdGVzLlJlYWRXcml0ZS5BbGwgVGFza3MuUmVhZFdyaXRlIFVzZXIuUmVhZCBVc2VyLlJlYWRCYXNpYy5BbGwgVXNlci5SZWFkV3JpdGUgVXNlci5SZWFkV3JpdGUuQWxsIiwic3ViIjoiRFkxcC1GVHppcmNSNnhYLTRTRmx1aWJEb0tqRURWWlpBY0tlSHRaQnNYOCIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6IjkzZjMzNTcxLTU1MGYtNDNjZi1iMDlmLWNkMzMxMzM4ZDA4NiIsInVuaXF1ZV9uYW1lIjoieXViby5zdW5AZHhjLmNvbSIsInVwbiI6Inl1Ym8uc3VuQGR4Yy5jb20iLCJ1dGkiOiI1d2NCb0FHMFJFaWVxUzNrNHZRZEFBIiwidmVyIjoiMS4wIiwid2lkcyI6WyJiNzlmYmY0ZC0zZWY5LTQ2ODktODE0My03NmIxOTRlODU1MDkiXSwieG1zX3N0Ijp7InN1YiI6IkZVNmF4bEZxUTVmTXVCdlY1Yzh5c2VyRzRYWFUzWnVKTEpXRkZOak1GazQifSwieG1zX3RjZHQiOjEzOTY2MTc0MjJ9.fd0laW6IgysbQMQ8pNmPZuiOexClSqV-EgyT7Cfh8sYIz29SQ09r_wH0O87Hk0nSO4EMQrWo-I5zdbsfwpf0JgPpKvLf0-izUqlsErH2fmwEvY52wBA59JlUp5Fs_CPEB0Qybr_ANJo_t4GogRkw3V18T9-D9ARZ21GDQbBHXkV-mfSLPODA1ETz1rOT44YOVZZ6U9LGrtF8qODf5y0NlT0HlUDmooxL52rCkfvWSDt9ZbtJxpTQekl9R9oHrKoJBJdSlL3eJUlqyCztKoZLS6uss0oXvE8f0TwuRrS6FiRLuWftHg34nnYJ1o83uC13tOcPicBUFT58GzrTF0hlbg'
            //                 callMSGraph(messageEndpoint, response.accessToken, (messages, endpoint) => {
            //                     channel.messages = messages.value;
            //                     channel.messages.map((d, i) => {
            //                         const replyEndpoint = endpoint + `/${d.id}/replies`;
            //                         callMSGraph(replyEndpoint, response.accessToken, (replies, endpoint) => {
            //                             d.replies = replies;
            //                         })
            //                     })
            //                 })
            //             })
            //         })
            //     });
                
            // })
        }).catch( error => {
            console.error(error);
        })
}

function getGuild() {
    const target = {
        'value': 
        [
            {
                'id': '529043f5-a154-4efa-96a2-8a49dedbe5db',
                'displayName': 'Architecture Guild',
                'description': 'Architecture Guild to increase knowledge sharing, professional learning and IP development.'
            }, 
            {
                'id': '99298122-9a85-4d85-9e34-ecef7b8c49af',
                'displayName': 'Guilds @ DXC',
                'description': 'Guilds @ DXC'
            },
            {
                'id': 'dd651bdf-b660-43e6-af20-59a1a318c015',
                'displayName': 'Technology Guilds@DXC',
                'description': 'This team is for developers to interact and learn programming languages by collaborating together. Created by EDGE Engineering Global SMEs Team.'
            }
        ]
    };
    getTokenPopup(teamsRequest)
        .then(response => {
            const endpoint = graphConfig.graphMeEndpoint + `/joinedTeams`;
            callMSGraph(endpoint, response.accessToken, (data, ep) => {
                for (const guild of target.value) {
                    if(data.value.find(d => d.id === guild.id))
                        guild.isMember = true;
                }
                updateUI(target, 'guilds');
            }, true);
        }).catch( error => {
            console.error(error);
        })

    // updateUI(target, 'guilds');

}

function getEvents(groupId) {
    getTokenPopup(teamsRequest)
        .then(response => {
            const dateTime = getNowformatted();
            const eventsEndpoint = graphConfig.graphBetaEndpoint + `/groups/${groupId}/events?$top=100&$select=subject,body,id,start,end,attendees`;
            // const eventsEndpoint = graphConfig.graphBetaEndpoint + `/groups/${groupId}/events?$top=100&$select=subject,body,id,start,end,attendees&$filter=start/dateTime ge '${dateTime}'`;
            callMSGraph(eventsEndpoint, response.accessToken, updateUI, true);
        }).catch( error => {
            console.error(error);
        })
    
}

function getChannels(groupId) {
    getTokenPopup(teamsRequest)
        .then(response => {
            const channelEndpoint = graphConfig.graphBetaEndpoint + `/teams/${groupId}/channels?$select=id,displayName`;
            callMSGraph(channelEndpoint, response.accessToken, updateUI, true);
        }).catch( error => {
            console.error(error);
        })

}

function getPosts(channelId, groupId) {
    getTokenPopup(teamsRequest)
        .then(response => {
            const endpoint = graphConfig.graphBetaEndpoint + `/teams/${groupId}/channels/${channelId}/messages?$top=100`;
            callMSGraph(endpoint, response.accessToken, updateUI);
        }).catch( error => {
            console.error(error);
        })
}

function addToGroup(group) {
    getTokenPopup(teamsRequest) 
        .then(response => {
            callMSGraph(graphConfig.graphMeEndpoint, response.accessToken, (data) => {
                const endpoint = graphConfig.graphEndpoint + `/groups/${group}/members/$ref`;
                const body = {
                    "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${data.id}`
                }
                postMSGraph(endpoint, response.accessToken, JSON.stringify(body), (data) => {
                    if(data == 'ok') {
                        getGuild();
                    }
                });
            });
        }).catch( error => {
            console.error(error);
        })

}

function addToCalendar(event, group) {
    getTokenPopup(calendarRequest)
        .then(response => {
            const currentAccount = myMSALObj.getAllAccounts()[0];
            const attendee = {
                "emailAddress": {
                    "address": currentAccount.username,
                    "name": currentAccount.name
                },
                "status": {
                    "response": "Accepted",
                    "time": `${getNowformatted()}Z`
                },
                "type": "optional"
            };
            let attendees = getFromCache(event).attendees;
            attendees.push(attendee);
            const body = {
                "attendees": attendees
            };

            const endpoint = graphConfig.graphEndpoint + `/groups/${group}/calendar/events/${event}`
            patchMSGraph(endpoint, response.accessToken, JSON.stringify(body), data => {
                if(data) 
                    alert(`${body.subject}\n\n This event has been successfully added to your calendar`);
            })
            // const endpoint = graphConfig.graphMeEndpoint + `/calendar/events`;
            
            // let body = getFromCache(event);
            // body.attendees = [];
            
            // postMSGraph(endpoint, response.accessToken, JSON.stringify(body), (data) => {
            //     if(data) {
            //         alert(`${body.subject}\n\n This event has been successfully added to your calendar`);
            //     }
            // });
        }).catch( error => {
            console.error(error);
        })
}

function getFromCache(id) {

    for(let i = 0; i < sessionStorage.length; i++) {
        if(sessionStorage.key(i).indexOf('graph.microsoft.com') > 0) {
            const value = sessionStorage.getItem(sessionStorage.key(i));
            const collection = JSON.parse(value);
            for (const item of collection) {
                if (item.id == id)
                    return item;
            }
        }
        
    }
    return;
}

function search() {
    const term = document.getElementById('search-term').value.toLowerCase();
    let result = [];
    for(let i = 0; i < sessionStorage.length; i++) {
        if(sessionStorage.key(i).indexOf('graph.microsoft.com') > 0) {
            const value = sessionStorage.getItem(sessionStorage.key(i));
            const collection = JSON.parse(value);
            for (const item of collection) {
                if (item.body.content) {
                    if(item.body.content.toLowerCase().indexOf(term) > 0) {
                        result = result.concat(item);
                    }
                }    
                
            }
            
        }
    }

    updateUI(result, 'search_result');
}

selectAccount();

function getNowformatted() {
    const date = new Date();
    const month = date.getMonth() + 1 > 9 ? date.getMonth() + 1 : `0${date.getMonth() + 1}`;
    const day = date.getDate() > 9 ? date.getDate() : `0${date.getDate()}`;

    const hour = date.getHours() > 9 ? date.getHours() : `0${date.getHours()}`;
    const minute = date.getMinutes() > 9 ? date.getMinutes() : `0${date.getMinutes()}`;
    const second = date.getSeconds() > 9 ? date.getSeconds() : `0${date.getSeconds()}`;
    return `${date.getFullYear()}-${month}-${day}T${hour}:${minute}:${second}`;
}

// export {signIn, signOut, seeProfile, readMail, getPerson, getTeams}