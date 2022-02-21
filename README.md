# webdriver
*Internet Explorer* が 2022/6/15 にサポート対象外になることにより、*windows script host* ではブラウザーの操作ができなくなると予想されます。

今後ブラウザを操作するならば、*WebDriver* 経由で *Microsoft Edge based on Chromium* を操作が主流になります。

このモジュールは *WebDriver* の操作を支援します。

## このモジュールのインストール

```shell
wes install @wachaon/webdriver --unsafe --bare
```

## *WebDriver* をインストールする

*WebDriver* とブラウザは同じバージョンのものを使用する必要があります。

このモジュールをコマンドラインから `-d` もしくは `--download` オプションを指定して実行すると、ブラウザのバージョンとアーキテクチャーが同じ *WebDriver* をダウンロードし展開します。

```shell
wes webdriver --download
```

## usage

下記スクリプトを実行して動作を確認してくだい。

```javascript
const { Window } = require('edge')
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
document.addCookie({
    name: name,
    value: genGUID()
})
console.log('getCookie(%O) // => %O', name, document.getCookie(name))
document.deleteCookie(name)

window.close('finished')
```

## `Window` クラス

ブラウザウインドウを操作するクラスになります。

### `document`

`Document` クラスの実体になります。

### `constructor(port, spec)`

`port` は *web driver* と通信するポートを指定します。既定値は `9515` ですが、`9515` ポートが使用済みの場合は使用可能なポートを探します。

`spec` は *web driver* のファイルパスを指定します。既定値は `msedgedriver.exe` です。相対パスで指定する場合はカレントディレクトリからの相対パスにします。

### `rect(prop)`

ウインドウの位置とサイズを変更します。
`prop` は `{x, y, width, height}` を指定します。 `x`、`y` はウインドウ位置、`width`、`height` はウインドウサイズを指定します。

### `navigate(url)`

`url` に指定した *URL* に移動します。

### `close(message)`

ブラウザの操作を終了します。
`message` を指定した場合は、終了後にコンソールへ `message` を表示します。

### `getURL()`

ブラウザの現在の *URL* を取得します。

### `back()`

ブラウザの履歴から一つ戻ります。

## `Document` クラス

*HTML Document* から必要な要素を指定するクラスになります。

簡素化を優先する為、メソッドは最低限の実装になります。

### `querySelectorAll(selector)`

*CSS Selector* で要素を検索します。対象がある場合の戻り値は要素の配列になります。

### `getTitle()`

タイトルを取得します。

### `getSource()`

*html* を取得します。

### `executeScript(script, args)`

`script` は関数オブジェクトを渡します。`script` は内部で `String(script)` で文字列に変換されるので、
関数外部の参照はできません。

```javascript
document.executeScript(
    function (text) {
        document.getElementsByTagName('input')[0].value = text
    },
    ["hello world"]
)
```

### `takeScreenShot()`

スクリーンショットを撮ります。戻り値は `Byte[]` になります。

### `getCookie(name)`

クッキーを取得します。`name` があればその名前のクッキーを取得、`name` がなければ全てのクッキーを取得します。

### `setCookie(object)`

クッキーを設定します。

### `deleteCookie(name)`

`name` のクッキーを削除します。

## `Element` クラス

### `querySelectorAll(selector)`

現在の要素の子孫ノードに対して *CSS Selector* で要素を検索します。対象がある場合の戻り値は要素の配列になります。

### `getAttribute (attribute)`

現在の要素の *element attribute* を取得します。

### `getProperty (property)`

現在の要素の *javascript property* を取得します。

### `click()`

現在の要素をクリックします。

### `setValue(text)`

現在の *input* などの要素に `text` を入力します。