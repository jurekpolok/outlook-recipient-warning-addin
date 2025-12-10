/* global Office */

Office.onReady(function() {
    // Office is ready
});

function onMessageSendHandler(event) {
    // Just allow the send immediately - test if event.completed works
    event.completed({ allowEvent: true });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
