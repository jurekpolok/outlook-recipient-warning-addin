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

/**
 * Check if an email address is external (not from internal domains)
 */
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

/**
 * Extract email address from recipient object
 */
function getEmailAddress(recipient) {
    return recipient.emailAddress || recipient.address || "";
}

/**
 * Handler for OnMessageSend event
 * This runs automatically when user clicks Send
 */
function onMessageSendHandler(event) {
    var item = Office.context.mailbox.item;

    // Get To recipients first
    item.to.getAsync(function(toResult) {
        if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
            return;
        }

        var toRecipients = toResult.value || [];

        // Get CC recipients
        item.cc.getAsync(function(ccResult) {
            if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
                event.completed({ allowEvent: true });
                return;
            }

            var ccRecipients = ccResult.value || [];
            var totalToCc = toRecipients.length + ccRecipients.length;

            // Check if threshold exceeded
            if (totalToCc > RECIPIENT_THRESHOLD) {
                // Count external recipients
                var allRecipients = toRecipients.concat(ccRecipients);
                var externalCount = 0;

                for (var i = 0; i < allRecipients.length; i++) {
                    if (isExternalEmail(getEmailAddress(allRecipients[i]))) {
                        externalCount++;
                    }
                }

                var warningMessage = "You are sending to " + totalToCc + " recipients in To/CC fields.";

                if (externalCount > 0) {
                    warningMessage = warningMessage + " " + externalCount + " of them are external.";
                }

                warningMessage = warningMessage + "\n\nConsider using BCC for external recipients to protect their email addresses from being disclosed to all recipients.";

                // Show dialog with warning - SoftBlock allows user to override
                event.completed({
                    allowEvent: false,
                    errorMessage: warningMessage
                });
            } else {
                // Under threshold, allow send
                event.completed({ allowEvent: true });
            }
        });
    });
}

// Register the event handler
// NOTE: Office.onReady is NOT called for event-based activation on Windows
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
