/* global Office */

Office.onReady(() => {
  Office.actions.associate("showThankYou", showThankYou);
});

function showThankYou(event) {
  const dialogUrl = new URL("./thankyou-dialog.html", window.location.href).toString();

  Office.context.ui.displayDialogAsync(dialogUrl, { height: 30, width: 20 }, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const dialog = result.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg.message === "close") {
          dialog.close();
        }
      });
    } else {
      showFallbackNotification();
    }

    if (event && typeof event.completed === "function") {
      event.completed();
    }
  });
}

function showFallbackNotification() {
  const item = Office.context?.mailbox?.item;

  if (item?.notificationMessages) {
    item.notificationMessages.addAsync("thankYouMessage", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Thank you! This test Outlook add-in is working.",
      icon: "Icon.16x16",
      persistent: false
    });
  }
}
