/* global Office */

// Configuration
const RECIPIENT_THRESHOLD = 5;
const INTERNAL_DOMAINS = [
    "bcc.no",
    "bcc.media",
    "brunstad.tv",
    "bccyep.no",
    "bcc-crew.com"
];

Office.onReady(() => {
    // Register the function with Office
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});

/**
 * Check if an email address is external (not from internal domains)
 */
function isExternalEmail(email) {
    if (!email) return false;
    const emailLower = email.toLowerCase();
    return !INTERNAL_DOMAINS.some(domain => emailLower.endsWith("@" + domain.toLowerCase()));
}

/**
 * Extract email address from recipient object
 */
function getEmailAddress(recipient) {
    return recipient.emailAddress || recipient.address || "";
}

/**
 * Get recipients from a recipient field asynchronously
 */
function getRecipientsAsync(recipientField) {
    return new Promise((resolve, reject) => {
        recipientField.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value || []);
            } else {
                reject(result.error);
            }
        });
    });
}

/**
 * Handler for OnMessageSend event
 * This runs automatically when user clicks Send
 */
async function onMessageSendHandler(event) {
    try {
        const item = Office.context.mailbox.item;

        // Get To and CC recipients
        const [toRecipients, ccRecipients] = await Promise.all([
            getRecipientsAsync(item.to),
            getRecipientsAsync(item.cc)
        ]);

        const totalToCc = toRecipients.length + ccRecipients.length;

        // Check if threshold exceeded
        if (totalToCc > RECIPIENT_THRESHOLD) {
            // Count external recipients
            const allToCcRecipients = [...toRecipients, ...ccRecipients];
            const externalCount = allToCcRecipients.filter(r =>
                isExternalEmail(getEmailAddress(r))
            ).length;

            let warningMessage = `You are sending to ${totalToCc} recipients in To/CC fields.`;

            if (externalCount > 0) {
                warningMessage += ` ${externalCount} of them are external.`;
            }

            warningMessage += `\n\nConsider using BCC for external recipients to protect their email addresses from being disclosed to all recipients.\n\nDo you want to send anyway?`;

            // Show dialog with warning - SoftBlock allows user to override
            event.completed({
                allowEvent: false,
                errorMessage: warningMessage
            });
        } else {
            // Under threshold, allow send
            event.completed({ allowEvent: true });
        }

    } catch (error) {
        console.error("Error in onMessageSendHandler:", error);
        // On error, allow send to not block the user
        event.completed({ allowEvent: true });
    }
}

// Make function globally available
globalThis.onMessageSendHandler = onMessageSendHandler;
