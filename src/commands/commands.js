/* global Office */

const notificationId = "smartMeetingRoomStatus";

function getNotificationOptions() {
  return {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Smart Meeting Room is now working on your request",
    icon: "Icon.16x16",
    persistent: false,
  };
}

function completeEvent(event) {
  event.completed({ allowEvent: true });
}

function showStatusMessage(event) {
  const options = getNotificationOptions();

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    notificationId,
    options,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        completeEvent(event);
      } else {
        Office.context.mailbox.item.notificationMessages.addAsync(notificationId, options, () =>
          completeEvent(event)
        );
      }
    }
  );
}

function onAppointmentSendHandler(event) {
  showStatusMessage(event);
}

Office.onReady(() => {
  Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
});

export { onAppointmentSendHandler };
