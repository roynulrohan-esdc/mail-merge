const messages = {
    ERROR: {
        tag: "ERROR",
        tagClass: "tag-error",
        containerClass: "row-error",
    },
    WARN: {
        tag: "WARN",
        tagClass: "tag-warn",
        containerClass: "row-warn",
    },
    INFO: {
        tag: "INFO",
        tagClass: "tag-info",
        containerClass: "row-info",
    }
}

//provide override for console
const console = {
    error: function (value) {
        createConsoleMessage(value, getStackTrace(), messages.ERROR.tag);
    },

    warn: function (value) {
        createConsoleMessage(value, getStackTrace(), messages.WARN.tag);
    },

    info: function (value) {
        createConsoleMessage(value, getStackTrace(), messages.INFO.tag);
    },

    log: function (value) {
        createConsoleMessage(value, getStackTrace());
    },

    show: function () {
        document.getElementById("console").style.bottom = "0px";
    },

    hide: function () {
        document.getElementById("console").style.bottom = -document.getElementById("console").offsetHeight + "px";
    },

    clear: function () {
        document.getElementById("console-container").innerHTML = ""
    },
};

document.addEventListener("DOMContentLoaded", function () {

    let dragBar = document.getElementById("console-drag");

    dragBar.addEventListener("mousedown", function (e) {
        document.addEventListener("mousemove", dragConsole);
    });

    document.addEventListener("mouseup", function () {
        document.removeEventListener("mousemove", dragConsole);
    });
})

function dragConsole(event) {
    let consoleElement = document.getElementById("console");
    let actionBar = document.getElementById("console-actionbar");

    let height = window.innerHeight - event.y;

    //keep console within screen
    if (height > actionBar.offsetHeight && height < window.innerHeight)
        consoleElement.style.height = height + "px";
}

function hideConsole() {
    console.hide()
}

function clearConsole() {
    console.clear()
}

window.addEventListener('keyup', function (event) {
    if (event.keyCode === 123) {
        var bottom = document.getElementById("console").style.bottom;
        document.getElementById("console").style.bottom = bottom === "0px" ? console.hide() : console.show();
    }
})

//Creates a stack trace to use for console debugging
function getStackTrace(event) {
    //Create and throw an error to get stack on IE9
    var e = new Error();
    try { throw e; }
    catch (e) { };

    var stack = e.stack.toString().split(/\r\n|\n/);
    var location = stack[stack.length - 1].split("/");

    var trace = location[location.length - 1];
    trace = trace.substring(0, trace.length - 1);

    return trace;
}

function formatMessage(message, recursive) {
    let container = document.createElement("div")


    if (message === null || message === undefined)
        container.innerHTML = message
    else {
        switch (typeof message) {
            case "object":
                container.appendChild(messageObject(message, recursive))
                break;
            case "string":
                container.appendChild(messageString(message))
                break;
            case "number":
                container.appendChild(messageNum(message))
                break;
            case "boolean":
                container.appendChild(messageBool(message))
                break;
            default:
                break
        }
    }

    return container
}

function messageString(message) {
    let p = document.createElement("p")

    p.className = "message-string"
    p.innerText = message

    return p
}

function messageNum(num) {
    let p = document.createElement("p")

    p.className = "message-number"
    p.innerText = num

    return p
}

function messageObject(object, recursive) {

    let container = document.createElement("div")
    container.className = "message-object"

    if (!recursive)
        container.className += " console-message-container"

    let keys = Object.keys(object)

    if (Array.isArray(object))
        container.innerText = " Array []"
    else
        container.innerText = " Object {}"

    for (let i = 0; i < keys.length; i++) {
        let div = document.createElement("div")

        div.style.display = "flex"


        let keyElement = document.createElement("div")
        let valueElement = document.createElement("div")

        keyElement.className = "message-object-key"
        keyElement.innerText = " " + keys[i] + ":"

        if (typeof object[keys[i]] !== "object")
            div.style.alignItems = "center"

        valueElement = formatMessage(object[keys[i]], true)

        div.appendChild(keyElement)
        div.appendChild(valueElement)

        container.appendChild(div)
    }

    return container
}

function messageBool(bool) {
    let p = document.createElement("p")

    p.className = "message-bool"
    p.innerText = bool

    return p
}

function createConsoleMessage(value, location, type) {
    let consoleContainer = document.getElementById("console-container");
    let consoleScroll = document.getElementById("console-scroll")

    let message = document.createElement("div");
    let messageType = document.createElement("p");
    let messageText = document.createElement("div");
    let messageTrace = document.createElement("p");

    message.className = "console-row "
    messageText.className = "console-text";
    messageTrace.className = "console-location";

    messageText.appendChild(formatMessage(value));
    messageTrace.innerText = location;

    if (type) {
        message.className += messages[type].containerClass;
        messageType.className = "console-tag " + messages[type].tagClass;
        messageType.innerText = messages[type].tag + ": ";
        message.appendChild(messageType);
    }

    message.appendChild(messageText);
    message.appendChild(messageTrace);
    consoleContainer.appendChild(message);

    message.offsetHeight = messageText.offsetHeight;

    consoleScroll.scrollTop = consoleScroll.scrollHeight;
}