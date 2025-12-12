/* global Office */

// Configuration
const RECIPIENT_THRESHOLD = 10;
const EXTERNAL_THRESHOLD = 5;
const DEBOUNCE_MS = 300;
const INTERNAL_DOMAINS = [
    "bcc.no",
    "bcc.media",
    "brunstad.tv",
    "bccyep.no",
    "bcc-crew.com"
];

// Debounce timer
let debounceTimer = null;

/**
 * Debounced version of checkRecipients to prevent excessive API calls
 */
function debouncedCheckRecipients() {
    if (debounceTimer) {
        clearTimeout(debounceTimer);
    }
    debounceTimer = setTimeout(function() {
        debounceTimer = null;
        checkRecipients();
    }, DEBOUNCE_MS);
}

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Initial check
        checkRecipients();

        // Set up event listener for recipient changes - updates automatically with debouncing
        if (Office.context.mailbox.item.addHandlerAsync) {
            Office.context.mailbox.item.addHandlerAsync(
                Office.EventType.RecipientsChanged,
                debouncedCheckRecipients
            );
        }
    }
});

/**
 * Check if an email address is external (not from internal domains)
 * Returns true for empty/invalid emails (treat as external for safety)
 * @param {string} email - Email address to check
 * @returns {boolean} - True if external
 */
function isExternalEmail(email) {
    if (!email) return true;
    const emailLower = email.toLowerCase().trim();
    if (emailLower.length === 0) return true;
    return !INTERNAL_DOMAINS.some(domain => emailLower.endsWith("@" + domain.toLowerCase()));
}

/**
 * Extract email address from recipient object
 * @param {Object} recipient - Recipient object
 * @returns {string} - Email address
 */
function getEmailAddress(recipient) {
    if (!recipient) return "";
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

        // Get all recipients (BCC may not be available in some Outlook versions)
        const toRecipients = await getRecipientsAsync(item.to);
        const ccRecipients = await getRecipientsAsync(item.cc);
        let bccRecipients = [];
        if (item.bcc) {
            try {
                bccRecipients = await getRecipientsAsync(item.bcc);
            } catch (e) {
                // BCC access not available, continue without it
            }
        }

        const toCount = toRecipients.length;
        const ccCount = ccRecipients.length;
        const bccCount = bccRecipients.length;
        const totalToCc = toCount + ccCount;

        // Count external recipients in To/CC and BCC
        const allToCcRecipients = [...toRecipients, ...ccRecipients];
        const externalInToCc = allToCcRecipients.filter(r => isExternalEmail(getEmailAddress(r))).length;
        const externalInBcc = bccRecipients.filter(r => isExternalEmail(getEmailAddress(r))).length;
        const totalExternal = externalInToCc + externalInBcc;

        // Update counts display
        document.getElementById("to-count").textContent = toCount;
        document.getElementById("cc-count").textContent = ccCount;
        document.getElementById("bcc-count").textContent = bccCount;
        document.getElementById("total-count").textContent = totalToCc;
        recipientDetails.classList.remove("hidden");

        // Check threshold: warn if EITHER:
        // 1. >10 recipients in To/CC AND at least 1 external ANYWHERE (To/CC/BCC)
        // 2. 5 or more external recipients in To/CC (regardless of total)
        const shouldWarn = (totalToCc > RECIPIENT_THRESHOLD && totalExternal > 0) || (externalInToCc >= EXTERNAL_THRESHOLD);

        if (shouldWarn) {
            // Show warning
            statusIcon.innerHTML = "&#9888;";
            statusIcon.className = "status-icon warning";

            let warningText;
            let notificationText;
            let warningBoxText;

            if (externalInToCc >= EXTERNAL_THRESHOLD) {
                warningText = `Warning: ${externalInToCc} external recipients in To/CC`;
                warningBoxText = `You have ${externalInToCc} external recipients in the To/CC fields. Their email addresses will be visible to all other recipients.`;
                notificationText = `You have ${externalInToCc} external recipients in To/CC. Consider using BCC for external recipients to protect their privacy.`;
            } else if (externalInBcc > 0 && externalInToCc === 0) {
                warningText = `Warning: ${totalToCc} addresses visible to external`;
                warningBoxText = `External recipients in BCC will see all ${totalToCc} email addresses in the To/CC fields.`;
                notificationText = `External recipients in BCC will see all ${totalToCc} email addresses in To/CC. Consider moving recipients to BCC to protect their addresses from external parties.`;
            } else {
                warningText = `Warning: ${totalToCc} recipients (${externalInToCc} external)`;
                warningBoxText = `You have ${totalToCc} recipients (${externalInToCc} external) in the To/CC fields. All email addresses will be visible to everyone.`;
                notificationText = `You have ${totalToCc} recipients (${externalInToCc} external) in To/CC. Consider using BCC for external recipients to protect their privacy.`;
            }

            statusMessage.textContent = warningText;
            document.getElementById("warning-text").textContent = warningBoxText;
            warningBox.classList.remove("hidden");

            // Show external recipients info (total across To/CC/BCC)
            document.getElementById("external-count").textContent = totalExternal;
            externalWarning.classList.remove("hidden");

            // Show notification
            showNotification("Privacy Warning", notificationText);
        } else {
            // All good - either under threshold or no external anywhere
            statusIcon.innerHTML = "&#10003;";
            statusIcon.className = "status-icon ok";
            if (totalToCc > RECIPIENT_THRESHOLD) {
                statusMessage.textContent = `${totalToCc} recipient(s) - all internal, OK`;
            } else {
                statusMessage.textContent = `${totalToCc} recipient(s) in To/CC - OK`;
            }
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
