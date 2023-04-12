

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {

  statusUpdate("icon16" , "Hello World!");
}

//function validateBody(event) {

//    Office.initialize = function () {
//        mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);

//    }

//}



//function checkBodyOnlyOnSendCallBack(asyncResults) {


//    console.log("Send Item");

//    asyncResults.asyncContext.completed({ allowEvent: true });

//};