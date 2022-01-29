const WShell = require('WScript.Shell')
const { eraseInLine, cursorHrAbs, brightGreen } = require('ansi')
const { download } = require('filesystem')
const { resolve, WorkingDirectory } = require('pathname')

const GET = 'GET'
const POST = 'POST'
const DELETE = 'DELETE'
const BOL = cursorHrAbs(1) // beginning of line
const ELEMENT_ID = 'element-6066-11e4-a52e-4f735466cecf'
const State = ['UNINITIALIZED', 'LOADING', 'LOADED', 'INTERACTIVE', 'COMPLETED']
const spiner = progress(['|', '/', '-', '\\', '|', '/', '-', '\\'])

class Window {
    constructor(port, spec) {
        const IServerXMLHTTPRequest2 = require('MSXML2.ServerXMLHTTP')

        port = port || findUnusedPort(9515)
        spec = spec || 'msedgedriver.exe'
        const driver = WShell.Exec(`${spec} port=${port}`)

        var { value } = request(
            IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${port}/session`,
            {
                capabilities: {}
            },
            'Initialize Session'
        )
        const { sessionId } = value

        this.port = port
        this.driver = driver
        this.sessionId = sessionId
        this.IServerXMLHTTPRequest2 = IServerXMLHTTPRequest2
        this.document = new Document(this)
    }
    rect(prop) {
        request(
            this.IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${this.port}/session/${this.sessionId}/window/rect`,
            prop,
            'Set Window Rect'
        )
    }
    navigate(url) {
        request(
            this.IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${this.port}/session/${this.sessionId}/url`,
            {
                url: url
            },
            'Navegate URL'
        )
    }
    close(message) {
        request(
            this.IServerXMLHTTPRequest2,
            DELETE,
            `http://localhost:${this.port}/session/${this.sessionId}/window`,
            null,
            'Close Window'
        )
        request(
            this.IServerXMLHTTPRequest2,
            DELETE,
            `http://localhost:${this.port}/session/${this.sessionId}`,
            null,
            'Delete Session'
        )
        this.driver.Terminate()
        if (message != null) console.log(message)
    }
    getURL() {
        const res = request(
            this.IServerXMLHTTPRequest2,
            GET,
            `http://localhost:${this.port}/session/${this.sessionId}/url`,
            null,
            'Get URL'
        )
        return res ? res.value : null
    }
    back() {
        request(
            this.IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${this.port}/session/${this.sessionId}/back`,
            {},
            'Back History'
        )
    }
    forward() {
        request(
            this.IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${this.port}/session/${this.sessionId}/forward`,
            {},
            'Forward History'
        )
    }
    getStatus() {
        const res = request(
            this.IServerXMLHTTPRequest2,
            GET,
            `http://localhost:${this.port}/status`,
            null,
            'Get Status'
        )
        if (res) return res.value
    }
    getCookie(name) {
        const url = `http://localhost:${this.port}/session/${this.sessionId}/cookie` + (name != null ? '/' + name : '')
        const res = request(this.IServerXMLHTTPRequest2, GET, url, null, 'Get Cookie')
        return res ? res.value : null
    }
    addCookie(cookie) {
        // cookie: {name: string, value: string, domain: string?, httpOnly: boolean?, path: string?, secure: boolean?}
        const parameter = { cookie: cookie }
        const res = request(
            this.IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${this.port}/session/${this.sessionId}/cookie`,
            parameter,
            'Add Cookie'
        )
        return res ? res.value : null
    }
    deleteCookie(name) {
        request(
            this.IServerXMLHTTPRequest2,
            DELETE,
            `http://localhost:${this.port}/session/${this.sessionId}/cookie/${name}`,
            null,
            'Delete Cookie'
        )
    }
}

class Document {
    constructor(window) {
        this.parentWindow = window
    }
    querySelectorAll(selector) {
        const window = this.parentWindow
        const res = request(
            window.IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${window.port}/session/${window.sessionId}/elements`,
            {
                using: 'css selector',
                value: selector
            },
            'Select Elements'
        )
        const elms = res != null ? res.value.map((val) => new Element(this, val[ELEMENT_ID])) : null
        return elms
    }
    getTitle() {
        const window = this.parentWindow
        return request(
            window.IServerXMLHTTPRequest2,
            GET,
            `http://localhost:${window.port}/session/${window.sessionId}/title`
        )
    }
}

class Element {
    constructor(document, elementId) {
        this.parentDocument = document
        this.elementId = elementId
    }
    querySelectorAll(selector) {
        const document = this.parentDocument
        const window = document.parentWindow
        const res = request(
            window.IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${window.port}/session/${window.sessionId}/element/${this.elementId}/elements`,
            {
                using: 'css selector',
                value: selector
            },
            'Select Elements'
        )
        const elms = res != null ? res.value.map((val) => new Element(document, val[ELEMENT_ID])) : null
        return elms
    }
    getAttribute(attribute) {
        const window = this.parentDocument.parentWindow
        const res = request(
            window.IServerXMLHTTPRequest2,
            GET,
            `http://localhost:${window.port}/session/${window.sessionId}/element/${this.elementId}/attribute/${attribute}`,
            null,
            'Get Attribute'
        )
        return res ? res.value : 'null'
    }
    getProperty(property) {
        const window = this.parentDocument.parentWindow
        const res = request(
            window.IServerXMLHTTPRequest2,
            GET,
            `http://localhost:${window.port}/session/${window.sessionId}/element/${this.elementId}/property/${property}`,
            null,
            'Get Property'
        )
        return res ? res.value : 'null'
    }
    click() {
        const window = this.parentDocument.parentWindow
        request(
            window.IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${window.port}/session/${window.sessionId}/element/${this.elementId}/click`,
            {},
            'Click Element'
        )
    }
    setValue(text) {
        const window = this.parentDocument.parentWindow
        request(
            window.IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${window.port}/session/${window.sessionId}/element/${this.elementId}/value`,
            {
                text
            },
            'Input Value'
        )
    }
}

function getWebDriver() {
    const version = WShell.exec(
        'powershell -Command Get-AppxPackage -Name Microsoft.MicrosoftEdge.* | foreach{$_.Version}'
    ).StdOut.ReadAll()

    const architecture =
        require('WScript.Shell').Environment('Process').Item('PROCESSOR_ARCHITECTURE') === 'x86' ? '32' : '64'

    const url = `https://msedgedriver.azureedge.net/${version}/edgedriver_win${architecture}.zip`

    download(url, resolve(WorkingDirectory, `edgedriver_win${architecture}.zip`))
    console.log('%sDownload of webdriver is complete!', brightGreen)
}

// util
function request(Server, method, url, parameter, processing = '', finished = '') {
    Server.open(method, url, true)
    if (method.toUpperCase === POST) Server.setRequestHeader('Content-Type', 'application/json')
    if (parameter != null) Server.send(JSON.stringify(parameter))
    else Server.send()

    while (State[Server.readyState] != 'COMPLETED') {
        console.print('%S%S %S%S', eraseInLine(0), processing, spiner(), BOL)
        WScript.Sleep(50)
    }
    console.print('%S%S', eraseInLine(0), finished)

    const res = Server.responseText
    return JSON.parse(res)
}

function progress(indicator) {
    let i = 0
    return function increment() {
        return indicator[i++ % indicator.length]
    }
}

function findUnusedPort(port) {
    const command = 'netstat -nao'
    const netstat = WShell.Exec(command)
    const res = netstat.StdOut.ReadAll()

    while (true) {
        port = port || parseInt(Math.random() * (65535 - 49152)) + 49152
        const exp = new RegExp('(TCP|UDP)\\s+\\d{1,3}.\\d{1,3}.\\d{1,3}.\\d{1,3}:' + port + '\\s')
        if (!exp.test(res)) break
        port = null
    }
    return port
}

// exports
module.exports = {
    Window,
    Document,
    Element,
    request,
    getWebDriver
}

// command line
if (wes.Modules[wes.main].path === __filename) getWebDriver()
