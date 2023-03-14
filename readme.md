# SeleniumVBA

&nbsp;&nbsp; [SeleniumVBA](https://github.com/GCuser99/SeleniumVBA) は VBA 用の [Selenium](https://www.selenium.dev/ja/documentation/) バインディングである。

&nbsp;&nbsp; `sample.xlsm` は、SeleniumVBA を用いて Google Chrome Browser で web スクレイピングを行う際プログラムのサンプルである。<br>
&nbsp;&nbsp; 以下では、`Sheet1.cls` のコードを示し、要点を解説する。web ページの遷移に関する典型的な操作は含まれているものと思われるので、自作する場合の参考とされたい。

```basic
     1	Private Sub CommandButton1_Click()
     2	
     3	    Dim driver As Object
     4	    Dim caps As Object
     5	    Dim keys As Object
     6	
     7	    Set driver = ThisWorkbook.Create_Object("WebDriver")
     8	    driver.StartChrome configure.driverPath
     9	    
    10	    If ThisWorkbook.use_proxy Then
    11	        Set caps = driver.CreateCapabilities
    12	        caps.AddArguments "--load-extension=" & configure.extension
    13	        driver.OpenBrowser caps
    14	    Else
    15	        driver.OpenBrowser
    16	    End If
    17	
    18	    driver.NavigateTo "https://www.selenium.dev"
    19	    driver.Wait 1000
    20	
    21	    driver.FindElementByClassName("DocSearch-Button").Click
    22	    driver.Wait 1000
    23	    
    24	    Set keys = ThisWorkbook.Create_Object("WebKeyboard")
    25	    driver.GetActiveElement().SendKeys "Getting started"
    26	    driver.Wait
    27	    driver.GetActiveElement().SendKeys keys.EnterKey
    28	    driver.Wait 5000
    29	
    30	    driver.CloseBrowser
    31	    driver.Shutdown
    32	
    33	End Sub
```
**1 行**&nbsp;&nbsp; `CommandButton1_Click` は「Selenium HP」ボタンをクリックした場合のイベント・プロシージャである。

![](img/sheet1.PNG)

**7 行**&nbsp;&nbsp; `WebDriver` クラスのインスタンスの生成。

![](img/navigate.PNG.PNG)

![](img/search.PNG)




![](img/sendKey.PNG)

**30 - 31 行**&nbsp;&nbsp; 後処理

&nbsp;&nbsp; 説明は割愛するが、このプログラムの前処理として `ThisWorkbook.cls` のサブルーチン `Workbook_open` で以下の処理を行っている。
自作する場合もこの。

* プロキシサーバーを経由する場合は、そのための情報を取得し、`Chrome WebDriver` のアドインを作成する。
* Google Chrome Browser　のバージョンに適合する `WebDriver` の最新版をインストールする。
* `SeleniumVBA` の最新版をインストールする。

## トラブルシュート

&nbsp;&nbsp; 起動時に「実行時エラー '6068': Visual Basic Project へのプログラム的なアクセスは信頼されません。」と表示される場合は、以下の手順でアクセスを明示的に許可する必要がある。

![](img/dev.PNG)

![](img/macroSetting.PNG)
