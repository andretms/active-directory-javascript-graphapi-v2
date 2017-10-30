// <sign-in>
// Initialize application
var userAgentApplication = new Msal.UserAgentApplication(msalconfig.clientID, null, loginCallback, {
    redirectUri: msalconfig.redirectUri
});

//Previous version of msal uses redirect url via a property
if (userAgentApplication.redirectUri) {
    userAgentApplication.redirectUri = msalconfig.redirectUri;
}

window.onload = function () {
    // If page is refreshed, continue to display user info
    if (!userAgentApplication.isCallback(window.location.hash) && window.parent === window && !window.opener) {
        updateUI();
    }
}

function updateUI() {
    var user = userAgentApplication.getUser();
    if (user) {

        // If user is already signed in, display the user info
        var userInfoElement = document.getElementById("userInfo");
        userInfoElement.parentElement.classList.remove("hidden");
        userInfoElement.innerHTML = JSON.stringify(user, null, 4);

        // Hide Sign-in button
        document.getElementById("signInButton").classList.add("hidden");

        // Show Query Graph API button
        document.getElementById("callGraphButton").classList.remove("hidden");

        // Show Sign-Out button
        document.getElementById("signOutButton").classList.remove("hidden");

    } else {

        // Show Sign-in button
        document.getElementById("signInButton").classList.remove("hidden");

        // Hide User Info
        document.getElementById("userInfo").parentElement.classList.add("hidden");

        // Hide Query Graph API button
        document.getElementById("callGraphButton").classList.add("hidden");

        // Hide Sign-Out button
        document.getElementById("signOutButton").classList.add("hidden");
    }
}

/**
 * Callback method from sign-in: if no errors, call callGraphApi() to show results.
 * @param {string} errorDesc - If error occur, the error message
 * @param {object} token - The token received from login
 * @param {object} error - The error 
 * @param {string} tokenType - the token type: usually id_token
 */
function loginCallback(errorDesc, token, error, tokenType) {
    if (errorDesc) {
        showError(msal.authority, error, errorDesc);
    } else {
        updateUI();
    }
}

function signIn() {
    var user = userAgentApplication.getUser();
    // If user is not signed in, then prompt user to sign in via loginRedirect.
    // This will redirect user to the Azure Active Directory v2 Endpoint
    if (!user) {
        userAgentApplication.loginPopup().then(function(idToken) {
            updateUI();
        });
    }
}

// </sign-in>

// <callgraph>
/**
 * Call the Microsoft Graph API and display the results on the page
 */
function callGraphApi() {
    var user = userAgentApplication.getUser();
    if (!user) {
        signIn().then(
            callGraphApi()
        );
    } else {
        // Graph API endpoint to show user profile
        var graphApiEndpoint = "https://graph.microsoft.com/v1.0/me";

        // Graph API scope used to obtain the access token to read user profile
        var graphApiScopes = ["https://graph.microsoft.com/user.read"];
        
        // Now Call Graph API to show the user profile information:
        var graphCallResponseElement = document.getElementById("graphResponse");
        graphCallResponseElement.parentElement.classList.remove("hidden");
        graphCallResponseElement.innerText = "Calling Graph ...";

        // In order to call the Graph API, an access token needs to be acquired.
        // Try to acquire the token used to Query Graph API silently first
        userAgentApplication.acquireTokenSilent(graphApiScopes)
            .then(function (token) {
                //After the access token is acquired, call the Web API, sending the acquired token
                callWebApiWithToken(graphApiEndpoint, token, graphCallResponseElement, document.getElementById("accessToken"));

            }, function (error) {
                // If the acquireTokenSilent() method fails, then acquire the token interactively via acquireTokenRedirect().
                // In this case, the browser will redirect user back to the Azure Active Directory v2 Endpoint so the user 
                // can re-type the current username and password and/ or give consent to new permissions your application is requesting.
                // After authentication/ authorization completes, this page will be reloaded again and callGraphApi() will be called.
                // Then, acquireTokenSilent will then acquire the token silently, the Graph API call results will be made and results will be displayed in the page.
                if (error) {
                    userAgentApplication.acquireTokenPopup(graphApiScopes).then(function(token) {
                        callWebApiWithToken(graphApiEndpoint, token, graphCallResponseElement, document.getElementById("accessToken"));
                    });
                };
            });

    }
}

/**
 * Show an error message in the page
 * @param {string} endpoint - the endpoint used for the error message
 * @param {string} error - the error string
 * @param {object} errorElement - the HTML element in the page to display the error
 */
function showError(endpoint, error, errorDesc) {
    var formattedError = JSON.stringify(error, null, 4);
    if (formattedError.length < 3) {
        formattedError = error;
    }
    document.getElementById("errorMessage").innerHTML = "An error has occurred:<br/>Endpoint: " + endpoint + "<br/>Error: " + formattedError + "<br/>" + errorDesc;
    console.error(error);
}
// </callgraph>

// <callwebapi>
/**
 * Call a Web API using an access token.
 * 
 * @param {any} endpoint - Web API endpoint
 * @param {any} token - Access token
 * @param {object} responseElement - HTML element used to display the results
 * @param {object} showTokenElement = HTML element used to display the RAW access token
 */
function callWebApiWithToken(endpoint, token, responseElement, showTokenElement) {
    var headers = new Headers();
    var bearer = "Bearer " + token;
    headers.append("Authorization", bearer);
    var options = {
        method: "GET",
        headers: headers
    };

    fetch(endpoint, options)
        .then(function (response) {
            var contentType = response.headers.get("content-type");
            if (response.status === 200 && contentType && contentType.indexOf("application/json") !== -1) {
                response.json()
                    .then(function (data) {
                        // Display response in the page
                        console.log(data);
                        responseElement.innerHTML = JSON.stringify(data, null, 4);
                        if (showTokenElement) {
                            showTokenElement.parentElement.classList.remove("hidden");
                            showTokenElement.innerHTML = token;
                        }
                    })
                    .catch(function (error) {
                        showError(endpoint, error);
                    });
            } else {
                response.json()
                    .then(function (data) {
                        // Display response as error in the page
                        showError(endpoint, data);
                    })
                    .catch(function (error) {
                        showError(endpoint, error);
                    });
            }
        })
        .catch(function (error) {
            showError(endpoint, error);
        });
}
// </callwebapi>

// <sign-out>
/**
 * Sign-out the user
 */
function signOut() {
    userAgentApplication.logout().then(function() {
        updateUI();
    });
}
// </sign-out>