Office.onReady(function() {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  item.getAttachmentsAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      event.completed({ allowEvent: true });
      return;
    }
    var attachments = result.value || [];
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
        errorMessage: "ISCOルールとして、ファイルの直接添付は原則禁止していますので、OneDrive等へアップロードしての送信をお願いします。やむを得ない場合に限り、「そのまま送信する」としてください。"
      });
    } else {
      event.completed({ allowEvent: true });
    }
  });
}
