export function getQueryParameter(paramName, searchStr) {
    // searchStr is from location.search
    let searchString = searchStr.substring(1),
        i, val, params = searchString.split("&");

    for (i = 0; i < params.length; i++) {
        val = params[i].split("=");
        if (val[0] == paramName) {
            return val[1];
        }
    }
    return null;
}
export function CustomException(message, internalException) {
    this.message = message;
    this.internal = internalException;
}
