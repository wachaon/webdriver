const WShell = require('WScript.Shell')
const { eraseInLine, cursorHrAbs, color } = require('ansi')
const { download } = require('filesystem')
const { resolve, WorkingDirectory, toPosixSep } = require('pathname')
const { unzip } = require('zip')
const { moveFileSync, deleteFileSync, deletedirSync, existsFileSync, existsdirSync } = require('filesystem')
const { has } = require('argv')

const GET = 'GET'
const POST = 'POST'
const DELETE = 'DELETE'
const BOL = cursorHrAbs(1) // beginning of line
const ELEMENT_ID = 'element-6066-11e4-a52e-4f735466cecf'
const State = ['UNINITIALIZED', 'LOADING', 'LOADED', 'INTERACTIVE', 'COMPLETED']
const spiner = progress(['|', '/', '-', '\\'])
const orange = color('#FFA500')

class Window {
    constructor(port, spec) {
        const IServerXMLHTTPRequest2 = require('MSXML2.ServerXMLHTTP')

        port = port || findUnusedPort(9515)
        spec = spec || 'msedgedriver.exe'
        const driver = WShell.Exec(`${spec} --port=${port} --silent`)

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
    executeScript(script = function () { }, args = []) {
        const code = `return (${String(script)})(...arguments)`
        const window = this.parentWindow
        const res = request(
            window.IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${window.port}/session/${window.sessionId}/execute/sync`,
            {
                script: code,
                args
            },
            'Execute Script'
        )
        return res ? res.value : null
    }
    xpath(path) {
        const window = this.parentWindow
        const res = request(
            window.IServerXMLHTTPRequest2,
            POST,
            `http://localhost:${window.port}/session/${window.sessionId}/elements`,
            {
                using: 'xpath',
                value: path
            },
            'Select Elements'
        )
        const elms = res || null
        return elms
    }
    getSource() {
        const window = this.parentWindow
        const res = request(
            window.IServerXMLHTTPRequest2,
            GET,
            `http://localhost:${window.port}/session/${window.sessionId}/source`,
            {},
            'Get Source'
        )
        return res ? res.value : null
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

// util
function getChromeVersion(spec = '"C:\\Program Files (x86)\\Google\\Chrome\\Application"') {
    const [version] = WShell.exec(
        `cmd /C dir /B /O-N ${spec}`
    ).StdOut.ReadAll().trim().split(/\r?\n/).slice(-1)
    return version
}

function getChromeDriverVersion(spec = 'chromedriver.exe') {
    return WShell.exec(
        `cmd /C ${spec} -v`
    ).StdOut.ReadAll().trim().replace(/^ChromeDriver ([\d\.]+) .+/, '$1')
}

function getFireFoxVersion(spec = '"C:\\Program Files\\Mozilla Firefox\\firefox.exe"') {
    return WShell.exec(
        `cmd /C ${spec} -v`
    ).StdOut.ReadAll().trim().slice('Mozilla Firefox '.length)
}

function getFireFoxDriverVersion(spec = 'geckodriver.exe') {
    return WShell.exec(
        `cmd /C ${spec} -V`
    ).StdOut.ReadAll().trim().split(/\r?\n/)[0].replace(/^geckodriver ([\d\.]+) .+/, '$1')
}

function getEdgeVersion() {
    return WShell.exec(
        'powershell -Command Get-AppxPackage -Name Microsoft.MicrosoftEdge.* | foreach{$_.Version}'
    ).StdOut.ReadAll().trim()
}

function getEdgeDriverVersion(spec = 'msedgedriver.exe') {
    return WShell.exec(
        `cmd /C ${spec} -v`
    ).StdOut.ReadAll().trim().replace(/^MSEdgeDriver ([\d\.]+) .+/, '$1')
}

function request(Server, method, url, parameter, processing, finished) {
    Server.open(method, url, true)
    if (method.toUpperCase === POST) Server.setRequestHeader('Content-Type', 'application/json')
    if (parameter != null) Server.send(JSON.stringify(parameter))
    else Server.send()

    let display
    while (State[Server.readyState] != 'COMPLETED') {
        display = `${BOL}${processing} ${spiner()}${eraseInLine(0)}`
        if (processing !== null) console.print(display)
        WScript.Sleep(50)
    }
    if (finished != null) display = `${BOL}${finished}${eraseInLine(0)}`
    else display = `${BOL}${eraseInLine(0)}`
    console.print(display)

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

function getEdgeWebDriver() {
    const version = getEdgeVersion()
    const driver = getEdgeDriverVersion()
    if (version === driver) return console.log('Both are installed with the correct version // => %O', version)

    const filename = 'msedgedriver.exe'
    const architecture = WShell
        .Environment('Process')
        .Item('PROCESSOR_ARCHITECTURE') === 'x86' ? '32' : '64'

    const url = `https://msedgedriver.azureedge.net/${version}/edgedriver_win${architecture}.zip`
    const zipSpec = resolve(WorkingDirectory, `edgedriver_win${architecture}.zip`)
    let dirSpec
    let fileSpec = resolve(WorkingDirectory, filename)
    try {
        console.log(download(url, zipSpec))
        console.log('unzip %O', dirSpec = toPosixSep(unzip(zipSpec)))
        if (existsFileSync(fileSpec)) deleteFileSync(fileSpec)
        console.log(moveFileSync(resolve(dirSpec, filename), fileSpec))
        console.log(deletedirSync(dirSpec))
        console.log(deleteFileSync(zipSpec))
        console.log('%SDownload of webdriver is complete. version: %S', orange, version)
    } catch (error) {
        console.log('%SFailed to download webdriver. version %S', orange, version)
        throw error
    } finally {
        if (existsFileSync(zipSpec)) console.log(deleteFileSync(zipSpec))
        if (existsdirSync(dirSpec)) console.log(deletedirSync(dirSpec))
    }
}

// exports
module.exports = {
    Window,
    Document,
    Element,
    request,
    getEdgeWebDriver,
}

// command line
if (wes.Modules[wes.main].path === __filename) {
    if (has('d') || has('download')) getEdgeWebDriver()
}
