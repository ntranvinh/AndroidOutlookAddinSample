import * as OfficeHelpers from '@microsoft/office-js-helpers';
import { OutlookHelper} from '../utils';
import { dialogUrl } from '../config'

//this is in percent
const confirmH = 18
const confirmW = 25
const successH = 21
const successW = 25

/* Render application after Office initializes */
Office.initialize = function () {
    console.log('Office onReady!!!')
    console.log(process.env.NODE_ENV)

    // do not remove this code
    if (OfficeHelpers.Authenticator.isAuthDialog()) {
        return
    }
};


async function action(event) {
    try{
        await OutlookHelper.showResult("step 1")
        const result1 = await OutlookHelper.showDialog(dialogUrl + "?dialogType=confirm", confirmH, confirmW)
        console.log(result1)
        await OutlookHelper.showResult("step 2")
        const result2 = await OutlookHelper.showDialog(dialogUrl + "?dialogType=success&&ms=success", successH, successW)
        console.log(result2)
        await OutlookHelper.showResult("finished!")
    }
    catch(ex){
        console.log(ex)
    }
    
    event.completed()

}

function getGlobal() {
    return (typeof self !== "undefined") ? self :
        (typeof window !== "undefined") ? window :
            (typeof global !== "undefined") ? global : undefined
}

const g = getGlobal()
g.action = action