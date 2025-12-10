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
 */
function onMessageSendHandler(event) {
    try {
        var item = Office.context.mailbox.item;

        if (!item) {
            event.completed({ allowEvent: true });
            return;
        }

        // Get To recipients
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

                    warningMessage = warningMessage + " Consider using BCC for privacy.";

                    event.completed({
                        allowEvent: false,
                        errorMessage: warningMessage
                    });
                } else {
                    event.completed({ allowEvent: true });
                }
            });
        });
    } catch (e) {
        event.completed({ allowEvent: true });
    }
}

// Register the event handler immediately
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
