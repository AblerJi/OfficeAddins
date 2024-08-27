"use strict";

Office.onReady(() => {
    // If needed, Office.js is ready to be called
});

var STORAGEITEM_KEY = "MyStorageItem";
var myEvent;
var item;

function onMessageComposeHandler(event) {
    item = Office.context.mailbox.item;
    myEvent = event;
    var msg ="";
    
    item.body.prependAsync(`<p style="color: green;">---Added by Test Web add-in Ex---</p>`, { coercionType: Office.CoercionType.Html }, (asyncResult)=> {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
        }else
        {
            console.log("Added content to the beginning of the body of the item.");
        }
        myEvent.completed();
    });
}

function onMessageReadWithCustomHeaderHandler(event) {
    item = Office.context.mailbox.item;
    myEvent = event;

    var now = new Date();
    var timeStamp = now.getMonth() + "/" + now.getDate() + "/" + now.getFullYear() + " " + now.getHours() + ":" + now.getMinutes() + ":" + now.getSeconds();

    console.log(`${timeStamp} [onMessageReadWithCustomHeaderHandler] called on [ + ${item.subject} + ] `);
    myEvent.completed(); 
}
function onMessageReadWithCustomAttachmentHandler(event){
    item = Office.context.mailbox.item;
    myEvent = event;

    var now = new Date();
    var timeStamp = now.getMonth() + "/" + now.getDate() + "/" + now.getFullYear() + " " + now.getHours() + ":" + now.getMinutes() + ":" + now.getSeconds();

    console.log(`${timeStamp} [onMessageReadWithCustomAttachmentHandler] called on [ + ${item.subject} + ] `);
    myEvent.completed();
}

async function checkSend(event){
    let details =
    {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "TestWebAddinInfo: Validating the email before allow sending. Please wait...",
        icon: "Icon.16x16",
        persistent: false
    };

    await Office.context.mailbox.item.notificationMessages.addAsync("TestWebAddinInfo", details);
    await my_sleep(2000);

    event.completed({ allowEvent: true });
}
const my_sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onMessageReadWithCustomHeaderHandler", onMessageReadWithCustomHeaderHandler);
Office.actions.associate("onMessageReadWithCustomAttachmentHandler", onMessageReadWithCustomAttachmentHandler);