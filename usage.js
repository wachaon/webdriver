const { Window } = require('./index')
const genGUID = require('genGUID')

const window = new Window()
const { document } = window
window.rect({
    x: 0,
    y: 0,
    width: 1280,
    height: 600
})
window.navigate('https://www.google.co.jp')

let [input] = document.querySelectorAll('[name="q"]')
input.setValue('selenium')

let [form] = document.querySelectorAll('form')
let [but] = form.querySelectorAll('input[name="btnK"]')
console.log('outerHTML // => %O', but.getProperty('outerHTML'))
console.log('but class=%O', but.getAttribute('class'))
but.click()

let status = window.getStatus()
console.log('Status: %O', status)

const name = genGUID()
window.addCookie({
    name: name,
    value: genGUID()
})
console.log('getCookie(%O) // => %O', name, window.getCookie(name))
window.deleteCookie(name)

window.close('finished')