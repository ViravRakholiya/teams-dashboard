// Create the main myMSALObj instance
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

    myMSALObj.loginPopup(loginRequest)
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

    const logoutRequest = {
        account: myMSALObj.getAccountByUsername(username),
        postLogoutRedirectUri: msalConfig.auth.redirectUri,
        mainWindowRedirectUri: msalConfig.auth.redirectUri
    };

    myMSALObj.logoutPopup(logoutRequest);
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

function seeUsers() {
    getTokenPopup(tokenRequest)
        .then(response => {
            callMSGraph(graphConfig.graphUsersEndpoint, response.accessToken, updateUI);
        }).catch(error => {
            console.error(error);
        });
}

function filterUsers(select,data){
    // Filter by team 
    if(select.value == 0){
        seeUsers()
    }else if(select.value == 1){
        filterbyTeam(data)
    }
}

function filterbyTeam(users){
    var teamcounter = 1
    if (users.value && users.value.length > 0)
        users.value.map((data, i) => {
            var userAdded = false
            getTokenPopup(tokenRequest)
                .then(response => {
                    var api = format(graphConfig.graphTeamEndpoint,[data.id])
                    callMSGraph(api, response.accessToken,function(response_data,endpoint){
                        if (response_data.value && response_data.value.length > 0){
                            response_data.value.forEach(servicePlans => {
                                servicePlans.servicePlans.forEach(servicePlan => {
                                    if(servicePlan["servicePlanName"] == "TEAMS1"){
                                        if(!userAdded){
                                            userAdded = true
                                            headerUiFOrFilter(teamcounter++,data)
                                            return
                                        }
                                    }
                                });
                            });
                        }
                    });
                }).catch(error => {
                    console.error(error);
                });
        })
}

function viewUsersStatus(type='normal',searchvalue=''){
    var userapiendpoint = graphConfig.graphUsersEndpoint;
    if(type=='search'){
        userapiendpoint = format(graphConfig.graphUserSearchEndPoint,[searchvalue]);
    }

    getTokenPopup(tokenRequest)
        .then(response => {
            let userIds = {"ids":[]}
            callMSGraph(userapiendpoint, response.accessToken, function(users,endpoint){
                if (users.value && users.value.length > 0) {
                    users.value.map((user,e)=>{
                        userIds["ids"].push(user.id+"")
                    });
                    if(userIds["ids"].length > 0){
                        getTokenPopup(tokenRequest)
                            .then(response2 => {
                                callMSGraphPost(userIds,graphConfig.graphUsersStatus,response2.accessToken,function(statusData,endpoint){
                                    var teamStatus = $.extend( true, statusData.value, users.value );
                                    updateUI(teamStatus,graphConfig.graphUsersStatus)
                                });
                            }).catch(error => {
                                console.error(error);
                            });
                    }
                }else{
                    alert('No user found');
                }
            });
        }).catch(error => {
            console.error(error);
        });
}

function seeUserSchedule(userId){
    var today = new Date();
    var startDate = today.toISOString();
    var endHours = today.getUTCHours() + 26;
    today.setHours(endHours)
    var endDate = today.toISOString();
    endDate = '2022-05-25T20:16:20.222Z'
    var api = format(graphConfig.graphUserEventsEndpoint,[userId,startDate,endDate]);
    getTokenPopup(tokenRequest)
        .then(response => {
            callMSGraph(api, response.accessToken, function(data,endpoint){
                displaySchedule(data.value);
            });
        }).catch(error => {
            console.error(error);
        });
}

function searchUser(searchValue){
    var api = format(graphConfig.graphUserSearchEndPoint,[searchValue]);
    console.log(api)
    getTokenPopup(tokenRequest)
        .then(response => {
            callMSGraph(api, response.accessToken, function(data,endpoint){
                updateUI(data,graphConfig.graphUsersEndpoint)
            });
        }).catch(error => {
            console.error(error);
        });
}

function viewGroups(){
    getTokenPopup(tokenRequest)
        .then(response => {
            callMSGraph(graphConfig.graphGetGroupEndPoint, response.accessToken, updateUI);
        }).catch(error => {
            console.error(error);
        });
}

function seeparticularGroup(groupID){
    var api = format(graphConfig.graphGetChannelEndPoint,[groupID])
    getTokenPopup(tokenRequest)
        .then(response => {
            callMSGraph(api, response.accessToken,function(data,e){
                console.log(data);
            });
        }).catch(error => {
            console.error(error);
        });
}

function format(source, params) {
    $.each(params,function (i, n) {
        source = source.replace(new RegExp("\\{" + i + "\\}", "g"), n);
    })
    return source;
}

function getDate(date){
    return date.getDate()+"."+(date.getMonth()+1)+"."+date.getFullYear();
}

function getTime(date){
    return date.getUTCHours()+":"+date.getUTCMinutes();
}

selectAccount();
