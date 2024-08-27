Office.onReady(function(info) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, CheckSelectedItem, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
        }
        console.log("Event handler added for the ItemChanged event.");
    });

    Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, CheckSelectedItem2, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
        }
        console.log("Event handler added for the SelectedItemsChanged event.");
    });

    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
    CheckSelectedItem();
});

function CheckSelectedItem() {
    if(Office.context != undefined && Office.context.mailbox != undefined && Office.context.mailbox.item != undefined)
    {
        console.log("Office.context.mailbox.item is not undefined");
        Office.context.mailbox.item.subject && $("#subject").text(Office.context.mailbox.item.subject);
        Office.context.mailbox.item.from && $("#from").text(`${Office.context.mailbox.item.from.displayName} (${Office.context.mailbox.item.from.emailAddress})`);
    }else
    {
        console.log("Office.context.mailbox.item is undefined, using getSelectedItemsAsync.");
        Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(asyncResult.error.message);
                return;
            }
            
            asyncResult.value.forEach((message) => {
                $("#subject").text(message.subject);
                console.log(`Item ID: ${message.itemId}`);
                console.log(`Subject: ${message.subject}`);
                console.log(`Item type: ${message.itemType}`);
                console.log(`Item mode: ${message.itemMode}`);
            });
        });
    }
}

function CheckSelectedItem2() {
    console.log("SelectedItemsChanged event fired.");
    Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
        }
        console.log("SelectedItems Info:");
        asyncResult.value.forEach((message) => {
            console.log(`Item ID: ${message.itemId}`);
            console.log(`Subject: ${message.subject}`);
            console.log(`Item type: ${message.itemType}`);
            console.log(`Item mode: ${message.itemMode}`);
        });
    });
}