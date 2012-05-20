// ==UserScript==
// @name Opera OWA
// @version 1.00
// @description Allows refreshing or opening Outlook in new tabs.
// @see https://gist.github.com/2699143/
// @include http*://outlook.com/*
// @include http*://*.outlook.com/*
// @include http*://owa.*/owa/*
// ==/UserScript==


window.addEventListener('load', function() {
    disableCookieCheck();
});


/**
 * Moves the `Cky` constructor functions to its prototype and disables the cookie-checking `IsLdd` function.
 * Prevents the following error when refreshing or opening Outlook in new tabs.
 *
 *     There was a problem opening your mailbox. You may have already signed in to Outlook Web App on a different browser tab. If so, close this tab and return to the other tab. If that doesn't work, you can try:
 *         - Closing your browser window and signing in again.
 *         - Deleting cookies from your browser and signing in again.
 *
 * The `createMasterWindowCookies` function uses the `Cky` method `IsLdd` to check the master window cookies and display the error.
 */
function disableCookieCheck()
{
    var RealCky = window.Cky;

    function Cky()
    {
        RealCky.apply(this, arguments);
        deleteFunctions(this);
        return this;
    }

    Cky.prototype = new RealCky();

    Cky.prototype.IsLdd = function Cky$IsLdd()
    {
        return false;
    };

    window.Cky = Cky;
}

function deleteFunctions(object)
{
    var name = null;
    for (name in object)
    {
        if (object.hasOwnProperty(name))
        {
            if (typeof object[name] === "function")
            {
                delete object[name];
            }
        }
    }
}
