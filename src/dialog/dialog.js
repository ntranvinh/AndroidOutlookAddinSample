import * as React from 'react';
import * as ReactDOM from "react-dom";
import {getQueryParameter} from "../utils"
import ConfirmOwa from "./components/confirm-owa";
import SuccessOwa from "./components/success-owa";

const urlQuery = location.search;
const ms = getQueryParameter("ms", urlQuery)

let loading = true

let dialogType = getQueryParameter("dialogType", urlQuery)

function loadUI() {
    const message = decodeURI(ms)

    let element = <h5>Dialog type, device type are missing or invalid</h5>
    if (dialogType != null) {
        if (dialogType === "success") {
            element = <SuccessOwa loading={loading}
                                      message={message}
                                      />
        }
        if (dialogType === "confirm") {
            element = <ConfirmOwa loading={loading}
                                      />
        }
    }
    ReactDOM.render(element, document.getElementById('container'))
}

/* Render application after Office initializes */
Office.onReady(() => {
    console.log('Dialog - Office onReady completely!!!')
    loading = false
    loadUI()
});

// load UI immediately,
// but disable some buttons, and checkboxes that need Office Js to finish loading
loadUI()