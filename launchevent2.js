Office.onReady(function () {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});

function onMessageSendHandler(event) {
  var done = false;
  function finish(options) {
    if (done) return;      // 二重呼び出し防止
    done = true;
    try { clearTimeout(watchdog); } catch (e) {}
    event.completed(options);
  }

  // 保険：5秒以内に判定できなければ、送信を止めずに通す（ハング防止）
  var watchdog = setTimeout(function () {
    finish({ allowEvent: true });
  }, 5000);

  try {
    var item = Office.context.mailbox.item;
    item.getAttachmentsAsync(function (result) {
      try {
        if (result.status === Office.AsyncResultStatus.Failed) {
          finish({ allowEvent: true });
          return;
        }
        var attachments = result.value || [];
        var hasFile = false;
        for (var i = 0; i < attachments.length; i++) {
          var att = attachments[i];
          if (att.attachmentType !== "cloud" &&
              att.attachmentType !== Office.MailboxEnums.AttachmentType.Cloud) {
            hasFile = true;
            break;
          }
        }
        if (hasFile) {
          finish({
            allowEvent: false,
            errorMessage: "ISCOルールとして、ファイルの直接添付は原則禁止していますので、OneDrive等へアップロードしての送信をお願いします。やむを得ない場合に限り、「そのまま送信する」としてください。"
          });
        } else {
          finish({ allowEvent: true });
        }
      } catch (e) {
        finish({ allowEvent: true });   // 予期せぬ例外でも固めない
      }
    });
  } catch (e) {
    finish({ allowEvent: true });
  }
}
