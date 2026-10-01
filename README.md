# mdb2CSV

`.mdb` ファイル内の全テーブルを、1テーブル1CSVで出力するツールです。  
UI はドラッグ&ドロップ対応（`tkinterdnd2` 利用時）です。

## 動作環境

- Windows
- Python 3.x
- `pyodbc`
- `tkinterdnd2`（未導入でもファイル選択UIで利用可能）
- Microsoft Access Database Engine (ODBC Driver)

## インストール

```powershell
cd mdb2CSV
py -m pip install -r requirements.txt
```

## 起動

```powershell
py mdb2CSV.py
```

または `mdb2CSV.py` をダブルクリックで起動できます。

## 使い方

1. `.mdb` ファイルをドラッグ&ドロップ、またはファイル選択で指定
2. 同じディレクトリに、`.mdb` と同名フォルダを作成
3. 全テーブルをCSV出力（UTF-8 BOM、ヘッダ行あり）
4. 必要に応じて「実行レポートを出力する」をONにすると、同じディレクトリに `<mdb名>_report.json` を追記出力

## 並び順（行順）の仕様

テーブル内レコードは、CSVに出力する文字列の昇順で並べます（数値も文字列として比較するため、`1, 10, 2` の順です）。

1. 主キー列を検出できた場合: 主キー列を優先し、同値なら全列で比較
2. 主キー未検出で unique index を検出できた場合: その列を優先し、同値なら全列で比較
3. 上記が未検出の場合: 全列を左から順に比較

NULLは従来どおり空文字として出力・比較します。同一行の重複は保持します。
Accessの照合順序や格納順に依存せず、同じ列構成・キー定義・内容なら同じCSVになります。
テーブルの処理順も名前順に固定します。

変更前と変更後のMDBは、両方ともこのバージョンでCSV化して比較してください。
ソートはテーブル単位でメモリ上で行うため、大きいテーブルではメモリ使用量が増えます。
実行レポートの全列ソート対象は `tables_sorted_by_all_columns` に記録します。

## 補足

- 主キー検出不可テーブルがある場合、完了メッセージに対象テーブル名を表示します。
- システムテーブル（`MSys*`, `USys*`, `~*`）は出力対象外です。

## よくあるエラー

- `No module named 'pyodbc'`
  - `py -m pip install -r requirements.txt` を実行
- MDB接続エラー
  - Microsoft Access Database Engine (ODBC Driver) の導入状況を確認
