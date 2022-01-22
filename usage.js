const {Window} = require('./index')

const window = new Window()
const {document} = window
window.rect({
    x: 0,
    y: 0,
    width: 1280,
    height: 600
})
window.navigate('https://www.google.co.jp')

let [input] = document.querySelectorAll('[name="q"]')
input.setValue('selenium')

let [but] = document.querySelectorAll('input[name="btnK"]')
but.click()

window.close('end')
