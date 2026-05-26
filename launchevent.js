console.log("[AttachmentWarning] launchevent.js loaded");

Office.onReady(function() {
  console.log("[AttachmentWarning] Office.onReady fired");
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  console.log("[AttachmentWarning] handler associated");
});

function onMessageSendHandler(event) {
  console.log("[AttachmentWarning] onMessageSendHandler called");
  var item = Office.context.mailbox.item;
  console.log("[AttachmentWarning] item:", item);
  
  item.getAttachmentsAsync(function(result) {
    console.log("[AttachmentWarning] getAttachmentsAsync result:", result.status, result.value);
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("[AttachmentWarning] failed, allowing event");
      event.completed({ allowEvent: true });
      return;
    }
    var attachments = result.value || [];
    var hasFile = false;
    for (var i = 0; i < attachments.length; i++) {
      var att = attachments[i];
      console.log("[AttachmentWarning] attachment", i, ":", att);
      if (att.attachmentType !== "cloud" && att.attachmentType !== Office.MailboxEnums.AttachmentType.Cloud) {
        hasFile = true;
        break;
      }
    }
    console.log("[AttachmentWarning] hasFile:", hasFile, "- calling event.completed");
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
