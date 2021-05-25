// Select DOM elements to work with
const welcomeDiv = document.getElementById("WelcomeMessage");
const signInButton = document.getElementById("SignIn");
const cardDiv = document.getElementById("card-div");
const mailButton = document.getElementById("readMail");
const profileButton = document.getElementById("seeProfile");
const profileDiv = document.getElementById("profile-div");

function showWelcomeMessage(username) {
    // Reconfiguring DOM elements
    cardDiv.style.display = 'initial';
    welcomeDiv.innerHTML = `Welcome ${username}`;
    signInButton.setAttribute("onclick", "signOut();");
    signInButton.setAttribute('class', "btn btn-success")
    signInButton.innerHTML = "Sign Out";
}

function updateUI(data, endpoint) {
    console.log('Graph API responded at: ' + new Date().toString());

    if (endpoint === graphConfig.graphMeEndpoint) {
        profileDiv.innerHTML = ''
        const title = document.createElement('p');
        title.innerHTML = "<strong>Title: </strong>" + data.jobTitle;
        const email = document.createElement('p');
        email.innerHTML = "<strong>Mail: </strong>" + data.mail;
        const phone = document.createElement('p');
        phone.innerHTML = "<strong>Phone: </strong>" + data.businessPhones[0];
        const address = document.createElement('p');
        address.innerHTML = "<strong>Location: </strong>" + data.officeLocation;
        profileDiv.appendChild(title);
        profileDiv.appendChild(email);
        profileDiv.appendChild(phone);
        profileDiv.appendChild(address);

    } else if (endpoint === graphConfig.graphMailEndpoint) {
        if (data.value.length < 1) {
            alert("Your mailbox is empty!")
        } else {
            const tabContent = document.getElementById("nav-tabContent");
            const tabList = document.getElementById("list-tab");
            tabList.innerHTML = ''; // clear tabList at each readMail call
            tabContent.innerHTML = '';

            data.value.map((d, i) => {
                // Keeping it simple
                if (i < 10) {
                    const listItem = document.createElement("a");
                    listItem.setAttribute("class", "list-group-item list-group-item-action")
                    listItem.setAttribute("id", "list" + i + "list")
                    listItem.setAttribute("data-toggle", "list")
                    listItem.setAttribute("href", "#list" + i)
                    listItem.setAttribute("role", "tab")
                    listItem.setAttribute("aria-controls", i)
                    listItem.innerHTML = d.subject;
                    tabList.appendChild(listItem)

                    const contentItem = document.createElement("div");
                    contentItem.setAttribute("class", "tab-pane fade")
                    contentItem.setAttribute("id", "list" + i)
                    contentItem.setAttribute("role", "tabpanel")
                    contentItem.setAttribute("aria-labelledby", "list" + i + "list")
                    contentItem.innerHTML = "<strong> from: " + d.from.emailAddress.address + "</strong><br><br>" + d.bodyPreview + "...";
                    tabContent.appendChild(contentItem);
                }
            });
        }
    } else if(endpoint.indexOf('people') > 0) {
        const tabContent = document.getElementById("nav-tabContent");
        const tabList = document.getElementById("list-tab");
        tabList.innerHTML = ''; 
        tabContent.innerHTML = '';

        data.value.map((d, i) => {
            const listItem = document.createElement("a");
            listItem.setAttribute("class", "list-group-item list-group-item-action")
            listItem.setAttribute("id", "list" + i + "list")
            listItem.setAttribute("data-toggle", "list")
            listItem.setAttribute("href", "#list" + i)
            listItem.setAttribute("role", "tab")
            listItem.setAttribute("aria-controls", i)
            listItem.innerHTML = d.displayName;
            tabList.appendChild(listItem)

            const contentItem = document.createElement("div");
            contentItem.setAttribute("class", "tab-pane fade")
            contentItem.setAttribute("id", "list" + i)
            contentItem.setAttribute("role", "tabpanel")
            contentItem.setAttribute("aria-labelledby", "list" + i + "list")
            contentItem.innerHTML = "<strong> Relevance score: " + d.scoredEmailAddresses[0].relevanceScore + "</strong><br><br>" +
                "<strong> Department: " + d.department + " </strong><br><br>" + 
                "<strong> Email: " + d.scoredEmailAddresses[0].address + " </strong><br><br>" +
                "<strong> Person type: " + d.personType.subclass + " </strong><br><br>"
            tabContent.appendChild(contentItem);
        });
    } else if(endpoint.indexOf('joinedTeams') > 0) {
        const tabContent = document.getElementById("nav-tabContent");
        const tabList = document.getElementById("list-tab");
        tabList.innerHTML = ''; 
        tabContent.innerHTML = '';

        data.value.map((d, i) => {
            // const listItem = document.createElement('div');
            // listItem.setAttribute("class", "list-group-item list-group-item-action")
            // listItem.setAttribute("id", d.id)
            // listItem.setAttribute("role", "tab")
            // listItem.setAttribute("aria-controls", i)
            // listItem.innerHTML = d.displayName;
            
            // const itemDetails = document.createElement('div');
            // itemDetails.innerHTML = d.description + 
            //     `<button onclick="getEvents('${d.id}')">events</button>` + 
            //     `<button onclick="getChannels('${d.id}')">channels</button>`

            const listItem = document.createElement("div");
            listItem.setAttribute("class", "card text-center")
            listItem.setAttribute("id", d.id)
            listItem.setAttribute("aria-labelledby", "list" + i + "list")
            const itemDetails = document.createElement("div");
            itemDetails.setAttribute("class", "card-body");
            itemDetails.setAttribute("style", "overflow: hidden");
            itemDetails.innerHTML = `<h5 class="card-title">${d.displayName}</h5><br><br>` +
                "<small>" + d.description + "</small><br><br>" + 
                `<button onclick="getEvents('${d.id}')">events</button>` + 
                `<button onclick="getChannels('${d.id}')">channels</button>`;

            listItem.appendChild(itemDetails)
            tabList.appendChild(listItem)
        })
    } else if(endpoint.indexOf('events') > 0) {
        const tabContent = document.getElementById("nav-tabContent");
        tabContent.innerHTML = '';

        data.value.map((d, i) => {
            const contentItem = document.createElement("div");
            contentItem.setAttribute("class", "card text-center")
            contentItem.setAttribute("id", "list" + i)
            contentItem.setAttribute("aria-labelledby", "list" + i + "list")
            const body = document.createElement("div");
            body.setAttribute("class", "card-body");
            body.setAttribute("style", "overflow: hidden");
            body.innerHTML = `<h5 class="card-title">${d.subject}</h5><br><br>` +
                "<small> Start: " + d.start.timeZone + " " + d.start.dateTime + " </small><br><br>" + 
                "<small> End: " + d.end.timeZone + " " + d.end.dateTime + " </small><br><br>" +
                "" + d.body.content + " <br><br>";
            contentItem.appendChild(body);
            tabContent.appendChild(contentItem);
        });
    } else if(endpoint.indexOf('channels') > 0 && endpoint.indexOf('messages') <0) {
        const teamId = data['@odata.context'].split(/'/)[1];
        const tabContent = document.getElementById(teamId);

        const channelList = document.getElementsByClassName('channel-list');
        const listLength = channelList.length;
        for(let i = 0; i < listLength; i++) {
            channelList[0].remove();
        }
  
        data.value.map((d, i) => {
            const channel = document.createElement("div");
            channel.setAttribute("class", "list-group-item list-group-item-action channel-list");
            channel.setAttribute("style", "overflow: hidden");
            channel.setAttribute("onclick", "getPosts('" + d.id + "', '" + teamId + "')")
            channel.innerHTML =  "<strong> " + d.displayName + " </strong><br><br>"
            // "<small> " + d.description + " </small>";

            tabContent.appendChild(channel);

        });

    } else if(endpoint.indexOf('messages') > 0) {
        const tabContent = document.getElementById("nav-tabContent");
        tabContent.innerHTML = '';

        data.value.map((d, i) => {
            const contentItem = document.createElement("div");
            contentItem.setAttribute("class", "card text-center")
            contentItem.setAttribute("id", "list" + i)
            contentItem.setAttribute("aria-labelledby", "list" + i + "list")
            const body = document.createElement("div");
            body.setAttribute("class", "card-body");
            body.setAttribute("style", "overflow: hidden");
            body.innerHTML = `` +
                "<small> Created: " + d.createdDateTime + " </small><br><br>" + 
                "<small> Last modified: " + d.lastModifiedDateTime + " </small><br><br>" +
                "" + d.body.content + " <br><br>";
            contentItem.appendChild(body);
            tabContent.appendChild(contentItem);
        });
    } else if(endpoint.indexOf('search_result') >= 0) {
        const modal = document.getElementById('search-modal');
        if(data.length == 0) {
            modal.innerHTML = 'No search result...'
            return ;
        }
        modal.innerHTML = '';

        data.map((d, i) => {
            const item = document.createElement("div");
            item.setAttribute("class", "list-group-item list-group-item-action channel-list");
            item.setAttribute("style", "overflow: hidden");
            if(d.start) {
                item.setAttribute("style", "background-color: aquamarine")
            }
            
            // item.setAttribute("onclick", "getPosts('" + d.id + "', '" + teamId + "')")
            item.innerHTML =  "<strong> " + d.body.content + " </strong><br><br>"
            // "<small> " + d.description + " </small>";

            modal.appendChild(item);
        });
    }
}
