Office.onReady(function () {
    // Set up onItemSend event handler
    Office.context.mailbox.item.addHandlerAsync(
        Office.EventType.ItemSend,
        onItemSendHandler
    );
});

function onItemSendHandler(eventArgs)
{
    // Get the meeting subject and organizer
    var subject = Office.context.mailbox.item.subject;
    var organizer = Office.context.mailbox.item.organizer.displayName;

    // Get the access token
    getAccessToken(function (accessToken) {
        // Create the CRM record
        createCrmRecord(subject, organizer, accessToken);
    });
}

function getAccessToken(callback) {
    // Configure MSAL.js
    const msalConfig = {
        auth: {
            clientId: 'clientid',
            authority: 'https://login.microsoftonline.com/tennatId',
        },
        cache: {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: true
        }
    };
    const msalInstance = new msal.PublicClientApplication(msalConfig);

    // Get the access token
    msalInstance.acquireTokenSilent({
        scopes: ['crmurl.default']
    }).then(function (authResult) {
        const accessToken = authResult.accessToken;
        callback(accessToken);
    }).catch(function (error) {
        console.error(error);
    });
}

function createCrmRecord(subject, organizer, accessToken) {
    // TODO: Implement your code to create the CRM record using the Web API.
    // For example:

    var endpointUrl = "crmurl/entityname";
    var payload = {
        "callsubject": subject,
        "name": organizer
    };
    var xhr = new XMLHttpRequest();
    xhr.open("POST", endpointUrl);
    xhr.setRequestHeader("Authorization", "Bearer " + accessToken);
    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    xhr.setRequestHeader("OData - MaxVersion", "4.0");
    xhr.setRequestHeader("OData-Version", "4.0");
    xhr.setRequestHeader("Prefer", "return=representation");

    xhr.onreadystatechange = function () {
        if (xhr.readyState === 4) {
            if (xhr.status === 204) {
                console.log("Record created successfully.");
            } else {
                console.log("Error creating record: " + xhr.statusText);
            }
        }
    };
    xhr.send(JSON.stringify(payload));
}



if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onAppointmentSendHandler", onItemSendHandler);
}