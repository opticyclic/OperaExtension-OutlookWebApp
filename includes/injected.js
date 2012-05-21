// ==UserScript==
// @name Opera OWA
// @version 1.00
// @description Allows refreshing or opening Outlook in new tabs. Fixes XML request bodies.
// @see https://gist.github.com/2699143/
// @include http*://outlook.com/*
// @include http*://*.outlook.com/*
// @include http*://owa.*/owa/*
// @include http*://*/CookieAuth.dll*
// ==/UserScript==

window.navigator.userAgent   = "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:15.0) Gecko/15.0 Firefox/15.0a1";
window.navigator.appVersion  = "5.0 (Windows)";

window.addEventListener('load', function() {
    disableCookieCheck();
    removeXmlDeclaration();
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

/**
 * Removes XML declarations from request bodies. Outlook servers reject request bodies with XML declarations.
 * Prevents the following error when sending requests:
 *
 *     An unexpected error occurred and your request couldn't be handled.
 */
function removeXmlDeclaration()
{
    var XmlDocumentPrototype = window.Owa.Dom.XmlDocument.prototype;
    var realGet_Xml = XmlDocumentPrototype.get_Xml;

    XmlDocumentPrototype.get_Xml = function get_Xml()
    {
        return realGet_Xml.call(this).replace(/^<\?xml\s+version\s*=\s*(["'])[^\1]+\1[^?]*\?>/i, "");
    };
}