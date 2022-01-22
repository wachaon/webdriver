# webdriver
*Internet Explorer* が 2022/6/16 から使用不可になることで、*windows script host* ではブラウザーの操作ができなくなります。

ブラウザーを操作するならば、*web driver* 経由で *Microsoft Edge based on Chromium* を操作する必要があります。

このモジュールは *web driver* の操作を簡略化します。

## このモジュールのインストール

```shell
wes install @wachaon/webdriver --unsafe --bare
```

## *web driver* をインストールする

*web driver* は使用するブラウザーとバージョンを同じにしなければなりません。

このモジュールをコマンドラインから直接指定すれば、ブラウザーのバージョンとアーキテクチャーが同じ *web driver* をダウンロードします。

```shell
wes webdriver
```

ダウンロードした zip を解凍して *msedgedriver.exe* を取り出してください。

## usage

下記スクリプトを実行して動作を確認してくだい。

```javascript
const {Window} = require('webdriver')

const window = new Window()
const {document} = window
window.rect({
    x: 0,
    y: 0,
    width: 1280,
    height: 600
})
window.navigate('https://www.google.com')

let [input] = document.querySelectorAll('[name="q"]')
input.setValue('selenium')

let [but] = document.querySelectorAll('input[name="btnK"]')
but.click()

window.close('end')
```

## `Window` クラス

ブラウザウインドウを操作するクラスになります。

### `document`

`Document` クラスの実体

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

## `Element` クラス

### `querySelectorAll(selector)`

現在の要素の子孫ノードに対して *CSS Selector* で要素を検索します。対象がある場合の戻り値は要素の配列になります。

### `getAttribute (attribute)`

現在の要素の *element attribute* を取得します。

### `getProperty (property)`

現在の要素の *javascript object* を取得します。

### `click()`

現在の要素をクリックします

### `setValue(text)`

現在の *input* などの要素に `text` を入力します。