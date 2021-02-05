/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. 
 *
 * This file shows how to use the SSO API to get a bootstrap token.
 */

 // If the add-in is running in Internet Explorer, the code must add support 
 // for Promises.
if (!window.Promise) {
    window.Promise = Office.Promise;
}

Office.onReady(function(info) {

    async function getGraphData() {
        try {
    
            let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    
            let exchangeResponse = await getGraphToken(bootstrapToken);
    
            if (exchangeResponse.claims) {
                let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
                exchangeResponse = await getGraphToken(mfaBootstrapToken);
            }
    
            if (exchangeResponse.error) {
                handleAADErrors(exchangeResponse);
            } 
            else {
                makeGraphApiCall(exchangeResponse.access_token);
            }
    
        }
        catch(exception) {
    
            if (exception.code) { 
                handleClientSideErrors(exception);
            }
            else {
                showMessage("EXCEPTION: " + JSON.stringify(exception));
            }
    
        }
    }

    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }

    function handleClientSideErrors(error) {
        switch (error.code) {
    
            case 13001:
                // No one is signed into Office. If the add-in cannot be effectively used when no one 
                // is logged into Office, then the first call of getAccessToken should pass the 
                // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
                // this error. 
                showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
                break;
            case 13002:
                // OfficeRuntime.auth.getAccessToken was called with the allowConsentPrompt 
                // option set to true. But, the user aborted the consent prompt. 
                showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
                break;
            case 13006:
                // Only seen in Office on the web.
                showMessage("Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
                break;
            case 13008:
                // The OfficeRuntime.auth.getAccessToken method has already been called and 
                // that call has not completed yet. Only seen in Office on the web.
                showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
                break;
            case 13010:
                // Only seen in Office on the web.
                showMessage("Follow the instructions to change your browser's zone configuration.");
                break;
                
            default:
                // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
                // and 50001, fall back to non-SSO sign-in.
                dialogFallback();
                break;
                
        }
    }

    function handleAADErrors(exchangeResponse) {

        if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1) && (retryGetAccessToken <= 0)) {
            retryGetAccessToken++;
            getGraphData();
        }
        
        else {
            dialogFallback();
        }
        
    }



    $(document).ready(function() {
        $('#getGraphDataButton').click(getGraphData);
    });
});

let retryGetAccessToken = 0;

