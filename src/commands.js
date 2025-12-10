/*
 * Recipient Privacy Warning - Outlook Add-in
 * Warns when sending to more than 5 external recipients
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
 * @param {string} email - Email address to check
 * @returns {boolean} - True if internal, false if external
 */
function isInternalEmail(email) {
    if (!email || email.length === 0) {
        return true; // Treat empty/null as internal (safe)
    }

    var emailLower = email.toLowerCase().trim();

    for (var i = 0; i < INTERNAL_DOMAINS.length; i++) {
        var domain = INTERNAL_DOMAINS[i].toLowerCase();
        // Check if email ends with @domain
        if (emailLower.indexOf("@" + domain) !== -1 &&
            emailLower.substring(emailLower.indexOf("@" + domain)) === "@" + domain) {
            return true;
        }
    }
    return false;
}

/**
 * Extract email address from recipient object
 * @param {Object} recipient - Recipient object from Office.js
 * @returns {string} - Email address
 */
function getEmailAddress(recipient) {
    if (!recipient) return "";
    return recipient.emailAddress || recipient.address || "";
}

/**
 * Main handler for OnMessageSend event
 * Called automatically when user clicks Send
 * @param {Object} event - Office.js event object
 */
function onMessageSendHandler(event) {
    var mailboxItem = Office.context.mailbox.item;

    // Get To recipients
    mailboxItem.to.getAsync(function(toResult) {
        if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
            // On error, allow send
            event.completed({ allowEvent: true });
            return;
        }

        var toRecipients = toResult.value || [];

        // Get CC recipients
        mailboxItem.cc.getAsync(function(ccResult) {
            if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
                // On error, allow send
                event.completed({ allowEvent: true });
                return;
            }

            var ccRecipients = ccResult.value || [];

            // Combine all recipients
            var allRecipients = toRecipients.concat(ccRecipients);

            // Count external recipients
            var externalCount = 0;
            var externalEmails = [];

            for (var i = 0; i < allRecipients.length; i++) {
                var email = getEmailAddress(allRecipients[i]);
                if (email && !isInternalEmail(email)) {
                    externalCount++;
                    externalEmails.push(email);
                }
            }

            // Check if external count exceeds threshold
            if (externalCount > RECIPIENT_THRESHOLD) {
                // Block and show warning
                event.completed({
                    allowEvent: false,
                    errorMessage: "You are sending to " + externalCount + " external recipients in To/CC fields. Consider moving some recipients to BCC to protect their email addresses from being shared with all recipients."
                });
            } else {
                // Allow send
                event.completed({ allowEvent: true });
            }
        });
    });
}

// IMPORTANT: Map the event handler name from manifest to JavaScript function
// This must be called at the top level, NOT inside Office.onReady or Office.initialize
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
