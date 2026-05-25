Office.onReady(function() {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});

function onMessageSendHandler(event) {
  Office.context.mailbox.item.attachments.getAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      event.completed({ allowEvent: true });
      return;
    }
    var attachments = result.value;
    var hasFile = false;
    for (var i = 0; i < attachments.length; i++) {
      var att = attachments[i];
      if (att.attachmentType !== "cloud" && att.attachmentType !== Office.MailboxEnums.AttachmentType.Cloud) {
        hasFile = true;
        break;
      }
    }
    if (hasFile) {
      event.completed({
        allowEvent: false,
        errorMessage: "データ添付ではなく、OneDrive等のリンクにて共有を検討してください"
      });
    } else {
      event.completed({ allowEvent: true });
    }
  });
}
