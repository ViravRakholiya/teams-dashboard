/** 
 * Helper function to call MS Graph API endpoint
 * using the authorization bearer token scheme
*/
function callMSGraph(endpoint, token, callback) {
    const headers = new Headers();
    const bearer = `Bearer ${token}`;

    headers.append("Authorization", bearer);
    headers.append("ConsistencyLevel", "eventual");

    const options = {
        method: "GET",
        headers: headers
    };

    console.log('request made to Graph API at: ' + new Date().toString());

    fetch(endpoint, options)
        .then(response => response.json())
        .then(response => callback(response, endpoint))
        .catch(error => console.log(error));
}

/** 
 * Helper function to call MS Graph API endpoint
 * using the authorization bearer token scheme
*/
function callMSGraphPost(requestbody, endpoint, token, callback) {
    const headers = new Headers();
    const bearer = `Bearer ${token}`;
    console.log(JSON.stringify(requestbody));
    headers.append("Authorization", bearer);
    headers.append("Content-Type","application/json")

    const options = {
        method: "POST",
        headers: headers,
        body: JSON.stringify(requestbody)
    };

    console.log('request made to Graph API at: ' + new Date().toString());
    
    fetch(endpoint, options)
        .then(response => response.json())
        .then(response => callback(response, endpoint))
        .catch(error => console.log(error));
}