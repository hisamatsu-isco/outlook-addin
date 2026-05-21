Office.onReady();

function onMessageSendHandler(event) {
  Office.context.mailbox.item.attachments.getAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Failed) {
      event.completed({ allowEvent: true });
      return;
    }

    var attachments = result.value;
    var hasLocalFile = false;

    for (var i = 0; i < attachments.length; i++) {
      var att = attachments[i];
      if (att.attachmentType !== "cloud" && att.attachmentType !== Office.MailboxEnums.AttachmentType.Cloud) {
        hasLocalFile = true;
        break;
      }
    }

    if (hasLocalFile) {
      event.completed({
        allowEvent: false,
        errorMessage: "ファイルが直接添付されています。\n\nOneDriveにアップロードしてリンクで共有することを推奨します。\n\nそのまま送信する場合は「とにかく送信」を押してください。"
      });
    } else {
      event.completed({ allowEvent: true });
    }
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
