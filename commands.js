/* global Office */

Office.onReady(() => {
  console.log('%c=== InboxAgent Commands Initialized ===', 'color: #0078d4; font-size: 14px; font-weight: bold;');
});

// Action button function
function action(event) {
  console.log('%c[COMMAND] Action function executed', 'color: #8b5cf6; font-weight: bold;');

  const item = Office.context.mailbox.item;

  if (item) {
    // Display a notification
    Office.context.mailbox.item.notificationMessages.addAsync(
      "actionNotification",
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "InboxAgent action executed successfully!",
        icon: "icon16",
        persistent: false
      },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log('✓ Notification displayed');
        }
      }
    );
  }

  // Signal that the function is complete
  event.completed();
}

// Register the function
Office.actions.associate("action", action);