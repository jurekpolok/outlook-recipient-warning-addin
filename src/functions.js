/* global Office */

// Configuration
const RECIPIENT_THRESHOLD = 5;
const INTERNAL_DOMAINS = [
    // Add your organization's internal email domains here
    "bcc.no",
    "bcc.media",
    "brunstad.tv",
    "bccyep.no",
    "bcc-crew.com"
];

Office.onReady(() => {
    // Register the event handler
});

/**
 * Event handler for when recipients change
 * This function is called automatically when recipients are added/removed
 * @param {Object} event - The event object
 */
function onRecipientsChanged(event) {
    checkRecipientsAndNotify(event);
}

/**
 * Check recipients and show notification if threshold exceeded
 * @param {Object} event - The event object
 */
async function checkRecipientsAndNotify(event) {
    try {
        const item = Office.context.mailbox.item;

        // Get all recipients
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

            let message = `You have ${totalToCc} recipients in To/CC fields.`;
            if (externalCount > 0) {
                message += ` ${externalCount} are external. Consider using BCC to protect their email addresses from being disclosed to all recipients.`;
            } else {
                message += ` Consider using BCC for privacy.`;
            }

            // Show notification message on the email
            if (item.notificationMessages) {
                item.notificationMessages.replaceAsync(
                    "recipientWarning",
                    {
                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                        message: message,
                        icon: "Icon.16x16",
                        persistent: true
                    }
                );
            }
        } else {
            // Remove warning if under threshold
            if (item.notificationMessages) {
                item.notificationMessages.removeAsync("recipientWarning");
            }
        }

    } catch (error) {
        console.error("Error in onRecipientsChanged:", error);
    }

    // Complete the event
    if (event && event.completed) {
        event.completed();
    }
}

/**
 * Check if an email address is external
 * @param {string} email - Email address to check
 * @returns {boolean} - True if external
 */
function isExternalEmail(email) {
    if (!email) return false;
    const emailLower = email.toLowerCase();
    return !INTERNAL_DOMAINS.some(domain => emailLower.endsWith("@" + domain.toLowerCase()));
}

/**
 * Extract email address from recipient object
 * @param {Object} recipient - Recipient object
 * @returns {string} - Email address
 */
function getEmailAddress(recipient) {
    return recipient.emailAddress || recipient.address || "";
}

/**
 * Get recipients from a recipient field asynchronously
 * @param {Object} recipientField - Office.js recipient field
 * @returns {Promise<Array>} - Array of recipients
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

// Register functions for Office
Office.actions = Office.actions || {};
Office.actions.associate("onRecipientsChanged", onRecipientsChanged);
