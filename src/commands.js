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

function isInternalEmail(email) {
    if (!email) return true;
    var emailLower = email.toLowerCase();
    for (var i = 0; i < INTERNAL_DOMAINS.length; i++) {
        if (emailLower.indexOf("@" + INTERNAL_DOMAINS[i].toLowerCase()) !== -1) {
            return true;
        }
    }
    return false;
}

function getEmailAddress(recipient) {
    return recipient.emailAddress || recipient.address || "";
}

function onMessageSendHandler(event) {
    var item = Office.context.mailbox.item;

    item.to.getAsync(function(toResult) {
        if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
            return;
        }

        var toRecipients = toResult.value || [];

        item.cc.getAsync(function(ccResult) {
            if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
                event.completed({ allowEvent: true });
                return;
            }

            var ccRecipients = ccResult.value || [];
            var allRecipients = toRecipients.concat(ccRecipients);

            // Count external recipients only
            var externalCount = 0;
            for (var i = 0; i < allRecipients.length; i++) {
                var email = getEmailAddress(allRecipients[i]);
                if (!isInternalEmail(email)) {
                    externalCount++;
                }
            }

            // Only warn if more than 5 external recipients
            if (externalCount > RECIPIENT_THRESHOLD) {
                event.completed({
                    allowEvent: false,
                    errorMessage: "You are sending to " + externalCount + " external recipients in To/CC. Consider using BCC to protect their email addresses from being disclosed to all recipients."
                });
            } else {
                event.completed({ allowEvent: true });
            }
        });
    });
}

// Register handler after Office is ready
Office.onReady(function() {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});
