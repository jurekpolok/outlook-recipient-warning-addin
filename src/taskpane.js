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

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Initial check
        checkRecipients();

        // Set up event listener for recipient changes - updates automatically
        if (Office.context.mailbox.item.addHandlerAsync) {
            Office.context.mailbox.item.addHandlerAsync(
                Office.EventType.RecipientsChanged,
                checkRecipients
            );
        }
    }
});

/**
 * Check if an email address is external (not from internal domains)
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
 * Main function to check recipients and display warnings
 */
async function checkRecipients() {
    const statusIcon = document.getElementById("status-icon");
    const statusMessage = document.getElementById("status-message");
    const recipientDetails = document.getElementById("recipient-details");
    const warningBox = document.getElementById("warning-box");
    const externalWarning = document.getElementById("external-warning");

    // Show loading state
    statusIcon.innerHTML = "&#8987;";
    statusIcon.className = "status-icon loading";
    statusMessage.textContent = "Checking recipients...";
    recipientDetails.classList.add("hidden");
    warningBox.classList.add("hidden");
    externalWarning.classList.add("hidden");

    try {
        const item = Office.context.mailbox.item;

        // Get all recipients
        const [toRecipients, ccRecipients, bccRecipients] = await Promise.all([
            getRecipientsAsync(item.to),
            getRecipientsAsync(item.cc),
            getRecipientsAsync(item.bcc)
        ]);

        const toCount = toRecipients.length;
        const ccCount = ccRecipients.length;
        const bccCount = bccRecipients.length;
        const totalToCc = toCount + ccCount;

        // Count external recipients in To and CC
        const allToCcRecipients = [...toRecipients, ...ccRecipients];
        const externalRecipients = allToCcRecipients.filter(r => isExternalEmail(getEmailAddress(r)));
        const externalCount = externalRecipients.length;

        // Update counts display
        document.getElementById("to-count").textContent = toCount;
        document.getElementById("cc-count").textContent = ccCount;
        document.getElementById("bcc-count").textContent = bccCount;
        document.getElementById("total-count").textContent = totalToCc;
        recipientDetails.classList.remove("hidden");

        // Check threshold
        if (totalToCc > RECIPIENT_THRESHOLD) {
            // Show warning
            statusIcon.innerHTML = "&#9888;";
            statusIcon.className = "status-icon warning";
            statusMessage.textContent = `Warning: ${totalToCc} recipients in To/CC fields`;
            warningBox.classList.remove("hidden");

            // Show external recipients info if any
            if (externalCount > 0) {
                document.getElementById("external-count").textContent = externalCount;
                externalWarning.classList.remove("hidden");
            }

            // Show notification
            showNotification(
                "Privacy Warning",
                `You have ${totalToCc} recipients in To/CC. Consider using BCC for external recipients to protect their privacy.`
            );
        } else {
            // All good
            statusIcon.innerHTML = "&#10003;";
            statusIcon.className = "status-icon ok";
            statusMessage.textContent = `${totalToCc} recipient(s) in To/CC - OK`;
        }

    } catch (error) {
        statusIcon.innerHTML = "&#10007;";
        statusIcon.className = "status-icon error";
        statusMessage.textContent = "Error checking recipients";
        console.error("Error checking recipients:", error);
    }
}

/**
 * Get recipients from a recipient field asynchronously
 * @param {Object} recipientField - Office.js recipient field (to, cc, or bcc)
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

/**
 * Show a notification to the user
 * @param {string} title - Notification title
 * @param {string} message - Notification message
 */
function showNotification(title, message) {
    // Use Office notification if available
    if (Office.context.mailbox.item.notificationMessages) {
        Office.context.mailbox.item.notificationMessages.replaceAsync(
            "recipientWarning",
            {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: message,
                icon: "Icon.16x16",
                persistent: true
            }
        );
    }
}
