
export class OutlookHelper {

    static getHost() {
        console.log(Office.context.mailbox.diagnostics.hostName)
        console.log(navigator)
        return Office.context.mailbox.diagnostics.hostName
    }

    static getCurrentEmailAddress() {
        return Office.context.mailbox.userProfile.emailAddress
    }

    static getCurrentDisplayName() {
        return Office.context.mailbox.userProfile.displayName
    }

    static async getCurrentEmailBody(type = Office.CoercionType.Html) {
        return new Promise((resolve, reject) => {
            Office.context.mailbox.item.body.getAsync(type,
                (result) => {
                    console.log(JSON.stringify(result));
                    resolve(result)
                }
            )
        })
    }

    static async getCurrentEmailBodyAsHtml() {
        const body = await this.getCurrentEmailBody(Office.CoercionType.Html)
        return body && body.value || null
    }

    static async getBoostrapGraphToken(options) {
        return new Promise((resolve, reject) => {
            Office.context.auth.getAccessTokenAsync(options,
                (result) => {
                    console.log(result)
                    resolve(result)
                }
            )
        })
    }


    static getCurrentEmailItemRestId() {
        if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
            // itemId is already REST-formatted
            return Office.context.mailbox.item.itemId;
        } else {
            // Convert to an item ID for API v2.0
            return Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.v2_0
            );
        }
    }


    static showRoamingStorage() {
        console.log(JSON.stringify(Office.context.roamingSettings))
    }

    static setRoamingStorage(name, value) {
        console.log("set roaming data: ", name)
        return new Promise((resolve, reject) => {
            Office.context.roamingSettings.set(name, value);
            Office.context.roamingSettings.saveAsync((asyncResult) => {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    // Handle the failure.
                    console.log(JSON.stringify(asyncResult))
                    reject(asyncResult)
                } else {
                    resolve(asyncResult)
                }
            })
        })
    }

    static getRoamingStorage(name) {
        const result = Office.context.roamingSettings.get(name)
        console.log("get roaming: ", name, !!result)
        return result
    }

    static setLocalStorage(name, value) {
        console.log("set local storage data: ", name, value)
        localStorage.setItem(name, JSON.stringify(value));
    }

    static getLocalStorage(name) {
        const result = localStorage.getItem(name);
        console.log("get local storage data: ", name, !!result)
        return JSON.parse(result)
    }

    static removeLocalStorage(name) {
        localStorage.removeItem(name);
        console.log("remove local storage data: ", name)
    }

    static deleteRoamingStorage(name) {
        console.log("remove roaming: ", name)
        return new Promise((resolve, reject) => {
            Office.context.roamingSettings.remove(name);
            Office.context.roamingSettings.saveAsync((asyncResult) => {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    // Handle the failure.
                    console.log(JSON.stringify(asyncResult))
                    reject(asyncResult)
                } else {
                    resolve(asyncResult)
                }
            })
        })

    }

    static async showResult(data, type='progressIndicator') {
        const key = "progress"
        await OutlookHelper.closeNotification(key)

        OutlookHelper.showNotification(key, data, type)
    }

    static showNotification(key, data, type){
        if (data !== '' && data.length > 150) {
            data = data.substr(0, 147) + '...'
        }
        
        console.log("show notification: ", data)

        //type can be errorMessage || progressIndicator || informationalMessage
        if (['errorMessage', 'progressIndicator', 'informationalMessage'].indexOf(type) < 0) {
            type = 'errorMessage'
        }

        Office.context.mailbox.item.notificationMessages.addAsync(key, {
            type,
            message: data
        });
    }

    static async showError(data, type='errorMessage'){
        const key = "error"
        await OutlookHelper.closeNotification(key)

        OutlookHelper.showNotification(key, data, type)
    }

    static closeNotification(key){
        if(['progress','error'].indexOf(key) < 0){
            key = 'progress'
        }
        return new Promise((resolve, reject) => {
            Office.context.mailbox.item.notificationMessages.removeAsync(key, function(a){
                resolve()
            });

        })
        
    }

    static cleanUpNotification(){
        return new Promise((resolve, reject) => {
            Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult) {
                
                if (asyncResult.value.length > 0) {
                    const cleanUpPromise = []
                    for(const value of asyncResult.value){
                        console.log(value)
                        cleanUpPromise.push(OutlookHelper.closeNotification(value.key))
                    }

                    return Promise.all(cleanUpPromise).then(function(){
                        resolve()
                    })
                }
            });

        })
        
    }

    static dialogCloseAsync(dialog, asyncResult) {
        // issue the close
        dialog.close();
        // and then try to add a handler
        // when that fails it is closed
        setTimeout(function () {
            try {
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, function () { });
                dialogCloseAsync(dialog, asyncResult);
            } catch (e) {
                asyncResult(); // done - closed
            }
        }, 1000);
    }

    static showDialog(url, height, width, logger=null) {
        let that = this
        return new Promise((resolve, reject) => {
            let dialog = null
            Office.context.ui.displayDialogAsync(url, { height: height, width: width, displayInIframe: true },
                function (asyncResult) {
                    // if(logger){
                    //     logger.captureMessage('dialog ' + url + " - return into callback: " + JSON.stringify(asyncResult));
                    // }
                    console.log('dialog ' + url + " - return into callback: " + JSON.stringify(asyncResult))

                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {            
                        reject(asyncResult);
                    } else {
                        dialog = asyncResult.value
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                            let result = JSON.parse(arg.message);
                            console.log(result)
                            //resolve(result)
                            that.dialogCloseAsync(dialog, () => {
                                // if(logger){
                                //     logger.captureMessage('dialog ' + url + " - function closed process completed");
                                // }
                                console.log('dialog ' + url + " - function closed process completed")
                                resolve(result)
                            })

                        })
                        dialog.addEventHandler(Office.EventType.DialogEventReceived, arg => {
                            let error = null
                            switch (arg.error) {
                                case 12002:
                                    error = "The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid."
                                    break;
                                case 12003:
                                    error = "The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required."
                                    break;
                                case 12006:
                                    error = "Dialog closed.";
                                    break;
                                default:
                                    error = "Unknown error in dialog box.";
                                    break;
                            }
                            //reject(new Error(`Dialog closed with error: ${error}`));
                            that.dialogCloseAsync(dialog, () => {
                                // if(logger){
                                //     logger.captureMessage('dialog' + url + "function closed process completed but dialog has error");
                                // }
                                console.log('dialog ' + url + "function closed process completed but dialog has error")
                                reject(new Error('Dialog closed with error: ' + error));
                            })
                        });
                    }
                }
            )
        })
    }

}