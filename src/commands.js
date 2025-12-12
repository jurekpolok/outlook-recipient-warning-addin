/*
 * Recipient Privacy Warning - Outlook Add-in v2.5.0.0
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
var TIMEOUT_MS = 10000;
var INTERNAL_DOMAINS = [
    "bcc.no",
    "bcc.media",
    "brunstad.tv",
    "bccyep.no",
    "bcc-crew.com"
];

/**
 * Check if email is from an internal domain
 * Returns false for empty/invalid emails (treat as external for safety)
 */
function isInternalEmail(email) {
    if (!email) {
        return false;
    }
    if (email.length === 0) {
        return false;
    }

    var emailLower = email.toLowerCase().trim();
    if (emailLower.length === 0) {
        return false;
    }

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
 * Count external recipients in an array
 */
function countExternalRecipients(recipients) {
    var count = 0;
    var i;
    var email;

    for (i = 0; i < recipients.length; i++) {
        email = getEmailAddress(recipients[i]);
        if (!isInternalEmail(email)) {
            count = count + 1;
        }
    }
    return count;
}

/**
 * Main handler for OnMessageSend event
 */
function onMessageSendHandler(event) {
    var mailboxItem = Office.context.mailbox.item;
    var eventCompleted = false;
    var timeoutId = null;

    /**
     * Safely complete the event (only once)
     */
    function completeEvent(options) {
        if (eventCompleted) {
            return;
        }
        eventCompleted = true;
        if (timeoutId !== null) {
            clearTimeout(timeoutId);
            timeoutId = null;
        }
        try {
            event.completed(options);
        } catch (e) {
            // Ignore errors if event already completed
        }
    }

    // Timeout wrapper - allow send after 10 seconds if something hangs
    timeoutId = setTimeout(function() {
        completeEvent({ allowEvent: true });
    }, TIMEOUT_MS);

    // Safety check for mailboxItem
    if (!mailboxItem) {
        completeEvent({ allowEvent: true });
        return;
    }

    // Get To recipients
    mailboxItem.to.getAsync(function(toResult) {
        var toRecipients;

        if (eventCompleted) {
            return;
        }

        if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
            completeEvent({ allowEvent: true });
            return;
        }

        if (toResult.value) {
            toRecipients = toResult.value;
        } else {
            toRecipients = [];
        }

        // Get CC recipients
        mailboxItem.cc.getAsync(function(ccResult) {
            var ccRecipients;

            if (eventCompleted) {
                return;
            }

            if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
                completeEvent({ allowEvent: true });
                return;
            }

            if (ccResult.value) {
                ccRecipients = ccResult.value;
            } else {
                ccRecipients = [];
            }

            // Get BCC recipients
            mailboxItem.bcc.getAsync(function(bccResult) {
                var bccRecipients;
                var toCcRecipients;
                var externalInToCc;
                var externalInBcc;
                var shouldWarn;
                var warningMessage;

                if (eventCompleted) {
                    return;
                }

                if (bccResult.status !== Office.AsyncResultStatus.Succeeded) {
                    completeEvent({ allowEvent: true });
                    return;
                }

                if (bccResult.value) {
                    bccRecipients = bccResult.value;
                } else {
                    bccRecipients = [];
                }

                // Combine To and CC for threshold check
                toCcRecipients = toRecipients.concat(ccRecipients);
                externalInToCc = countExternalRecipients(toCcRecipients);
                externalInBcc = countExternalRecipients(bccRecipients);
                var totalExternal = externalInToCc + externalInBcc;

                // Prompt if EITHER condition is met:
                // 1. Total To/CC recipients > 10 AND at least one external ANYWHERE (To/CC/BCC)
                // 2. 5 or more external recipients in To/CC (regardless of total)
                shouldWarn = false;
                warningMessage = "";

                if (externalInToCc >= EXTERNAL_THRESHOLD) {
                    shouldWarn = true;
                    warningMessage = "You are sending to " + externalInToCc + " external recipients in To/CC fields. Consider moving external recipients to BCC to protect their email addresses from being shared with all recipients.";
                } else if (toCcRecipients.length > RECIPIENT_THRESHOLD && totalExternal > 0) {
                    shouldWarn = true;
                    if (externalInBcc > 0 && externalInToCc === 0) {
                        warningMessage = "External recipients in BCC will see all " + toCcRecipients.length + " email addresses in To/CC. Consider moving recipients to BCC to protect their addresses from external parties.";
                    } else {
                        warningMessage = "You are sending to " + toCcRecipients.length + " recipients (" + externalInToCc + " external) in To/CC fields. Consider moving some recipients to BCC to protect their email addresses from being shared with all recipients.";
                    }
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
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
