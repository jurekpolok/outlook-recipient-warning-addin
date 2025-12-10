/* global Office */

// Configuration
var RECIPIENT_THRESHOLD = 5;
var INTERNAL_DOMAINS = [
    "bcc.no",
    "bcc.media",
    "brunstad.tv",
    "bccyep.no",
    "bcc-crew.com"
];

function isExternalEmail(email) {
    if (!email) return false;
    var emailLower = email.toLowerCase();
    for (var i = 0; i < INTERNAL_DOMAINS.length; i++) {
        if (emailLower.indexOf("@" + INTERNAL_DOMAINS[i].toLowerCase()) !== -1) {
            return false;
        }
    }
    return true;
}

function getEmailAddress(recipient) {
    return recipient.emailAddress || recipient.address || "";
}

function onMessageSendHandler(event) {
    // Simple test - just show warning for any email
    event.completed({
        allowEvent: false,
        errorMessage: "Test: Add-in is working! You have recipients in your email."
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
