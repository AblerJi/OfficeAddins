Office.onReady(function(info) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, CheckSelectedItem, (asyncResult) => {
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
        Office.context.mailbox.item.subject.getAsync((result)=>{
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                $("#subject").text(result.value);
            } else {
                console.log("Failed to get 'subject' data. Error: " + result.error.message);
            }
        });
        $("#recipients li").remove();
        Office.context.mailbox.item.to.getAsync((result)=>{
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const msgTo = result.value;
                if(msgTo.length > 0)
                    $("#recipients").append(`<li>To:</li>`);
                for (let i = 0; i < msgTo.length; i++) {
                    $("#recipients").append(`<li>${msgTo[i].displayName + " (" + msgTo[i].emailAddress + ")"}</li>`);
                }
            } else {
                console.log("Failed to get 'to' data. Error: " + result.error.message);
            }
        });

        Office.context.mailbox.item.cc.getAsync((result)=>{
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const msgCC = result.value;
                if(msgCC.length > 0)
                    $("#recipients").append(`<li>CC:</li>`);
                for (let i = 0; i < msgCC.length; i++) {
                    $("#recipients").append(`<li>${msgCC[i].displayName + " (" + msgCC[i].emailAddress + ")"}</li>`);
                }
            } else {
                console.log("Failed to get 'cc' data. Error: " + result.error.message);
            }
        });
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