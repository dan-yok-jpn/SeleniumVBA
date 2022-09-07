### WIS からのデータの取得

近年、Web スクレイピング用の標準ツールとなっている [Selenium](https://www.selenium.dev/ja/) を用いて国土交通省 水文水質データベース (Water Information System) からデータを取得する<sup>※</sup>。

<span style="font-size: small">
※ Excel のアプリケーション wis.xlsm は内部的に InternetExplorer (IE) を操作するためのクラスモジュールを用いている。IE のサポートは既に終了しており、これに替わる VBA のクラスモジュールの提供もアナウンスされていないため、OS に IE がバンドルされなくなると wis.xlsm は使用できなくなる。
</span>

#### 仮想環境の設定

Selenium は web ブラウザとの相性があり、常用のモジュールではないので仮想環境を設けてこれを使用する (用済みの場合は .venv ごと削除すれば良い)。
Chrome でスクレイピングを行う仮想環境はコマンドプロンプトで以下を実行すると作成される (同時に VSCode で編集・実行するための設定も行われる)。

```sh
% set_venv
```

ただし、以下のようなメッセージが表示される場合は set_venv.bat の PYTHONHOME を自身の環境に合わせて訂正する必要がある。

```sh
ERROR !   C:\Python\Python310 not found.
Check this Scripts
```

#### テスト

テスト用のモジュール wis.py は大町ダムの 2022/8/19 から 2022/8/25 のダム諸量を contents.csv (utf-8) に出力する単純なものであり、左記の条件は以下のようにハードコーディングされている。

```python
ID = "1368040365050" # ohmachi dam
DT1 = "20220819"
DT2 = "20220825"
```

実行に際しては仮想環境をアクティベートする必要がある。

```sh
% .venv\Scripts\activate
(.venv)% python wis.py
(.venv)% deactivate
``` 

#### 拡張について

Selenium を使用する処理については wis.py でほぼ実装されている (データを anchor 要素から取得するケースと iframe 要素から取得するケースを示し、簡潔な前者を採用した)。
対象となる観測局の指定やデータの種別・取得期間を与える前処理や取得したデータの後処理は別途のモジュールを記述して拡張すれば良い。
