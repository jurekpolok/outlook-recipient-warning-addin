/*
 * Recipient Privacy Warning - Outlook Add-in v2.5.0
 * Warns when sending to more than 10 recipients AND at least one is external
 * ECMAScript 2016 compatible - no ternary, no ||, no async/await
 */

// Call Office.onReady to satisfy the Office.js requirement
Office.onReady(function() {
    // Intentionally empty - event handlers don't need initialization here
});

// Configuration
var RECIPIENT_THRESHOLD = 10;
var EXTERNAL_THRESHOLD = 5;
var TIMEOUT_MS = 5000;
var INTERNAL_DOMAINS = [
    "bcc.no",
    "bcc.media",
    "brunstad.tv",
    "bccyep.no",
    "bcc-crew.com"
];

/**
 * Check if email is from an internal domain
 */
function isInternalEmail(email) {
    if (!email) {
        return true;
    }
    if (email.length === 0) {
        return true;
    }

    var emailLower = email.toLowerCase().trim();
    var i;
    var domain;
    var atDomain;

    for (i = 0; i < INTERNAL_DOMAINS.length; i++) {
        domain = INTERNAL_DOMAINS[i].toLowerCase();
        atDomain = "@" + domain;
        if (emailLower.indexOf(atDomain) !== -1) {
            if (emailLower.substring(emailLower.length - atDomain.length) === atDomain) {
                return true;
            }
        }
    }
    return false;
}

/**
 * Extract email address from recipient object
 */
function getEmailAddress(recipient) {
    if (!recipient) {
        return "";
    }
    if (recipient.emailAddress) {
        return recipient.emailAddress;
    }
    if (recipient.address) {
        return recipient.address;
    }
    return "";
}

/**
 * Main handler for OnMessageSend event
 */
function onMessageSendHandler(event) {
    var mailboxItem = Office.context.mailbox.item;
    var eventCompleted = false;

    // Timeout wrapper - allow send after 5 seconds if something hangs
    var timeoutId = setTimeout(function() {
        if (!eventCompleted) {
            eventCompleted = true;
            event.completed({ allowEvent: true });
        }
    }, TIMEOUT_MS);

    /**
     * Safely complete the event (only once)
     */
    function completeEvent(options) {
        if (!eventCompleted) {
            eventCompleted = true;
            clearTimeout(timeoutId);
            event.completed(options);
        }
    }

    // Safety check for mailboxItem
    if (!mailboxItem) {
        completeEvent({ allowEvent: true });
        return;
    }

    mailboxItem.to.getAsync(function(toResult) {
        var toRecipients;

        if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
            completeEvent({ allowEvent: true });
            return;
        }

        if (toResult.value) {
            toRecipients = toResult.value;
        } else {
            toRecipients = [];
        }

        mailboxItem.cc.getAsync(function(ccResult) {
            var ccRecipients;
            var allRecipients;
            var externalCount;
            var i;
            var email;

            if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
                completeEvent({ allowEvent: true });
                return;
            }

            if (ccResult.value) {
                ccRecipients = ccResult.value;
            } else {
                ccRecipients = [];
            }

            allRecipients = toRecipients.concat(ccRecipients);
            externalCount = 0;

            for (i = 0; i < allRecipients.length; i++) {
                email = getEmailAddress(allRecipients[i]);
                if (email) {
                    if (!isInternalEmail(email)) {
                        externalCount = externalCount + 1;
                    }
                }
            }

            // Prompt if EITHER condition is met:
            // 1. Total recipients > 10 AND at least one external recipient
            // 2. 5 or more external recipients (regardless of total)
            var shouldWarn = false;
            var warningMessage = "";

            if (externalCount >= EXTERNAL_THRESHOLD) {
                shouldWarn = true;
                warningMessage = "You are sending to " + externalCount + " external recipients in To/CC fields. Consider moving external recipients to BCC to protect their email addresses from being shared with all recipients.";
            } else if (allRecipients.length > RECIPIENT_THRESHOLD && externalCount > 0) {
                shouldWarn = true;
                warningMessage = "You are sending to " + allRecipients.length + " recipients (" + externalCount + " external) in To/CC fields. Consider moving some recipients to BCC to protect their email addresses from being shared with all recipients.";
            }

            if (shouldWarn) {
                completeEvent({
                    allowEvent: false,
                    errorMessage: warningMessage
                });
            } else {
                completeEvent({ allowEvent: true });
            }
        });
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
