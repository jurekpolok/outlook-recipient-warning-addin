/*
 * Recipient Privacy Warning - Outlook Add-in v2.1.0
 * Warns when sending to more than 5 external recipients
 * ECMAScript 2016 compatible - no ternary, no ||, no async/await
 */

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

    mailboxItem.to.getAsync(function(toResult) {
        var toRecipients;

        if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
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
                event.completed({ allowEvent: true });
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

            if (externalCount > RECIPIENT_THRESHOLD) {
                event.completed({
                    allowEvent: false,
                    errorMessage: "You are sending to " + externalCount + " external recipients in To/CC fields. Consider moving some recipients to BCC to protect their email addresses from being shared with all recipients."
                });
            } else {
                event.completed({ allowEvent: true });
            }
        });
    });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
