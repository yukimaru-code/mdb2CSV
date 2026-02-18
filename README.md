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

テーブル内レコードの出力順は、次の優先順で決定します。

1. 主キー列を検出できた場合: 主キー順で `ORDER BY`
2. 主キー未検出で unique index を検出できた場合: その列順で `ORDER BY`
3. 上記が未検出の場合: 先頭列で `ORDER BY`
4. 列情報も取れない場合: `ORDER BY` なし（DB返却順）

注: Access画面の見え方（保存済み並び替え等）と完全一致しない場合があります。

## 補足

- 主キー検出不可テーブルがある場合、完了メッセージに対象テーブル名を表示します。
- システムテーブル（`MSys*`, `USys*`, `~*`）は出力対象外です。

## よくあるエラー

- `No module named 'pyodbc'`
  - `py -m pip install -r requirements.txt` を実行
- MDB接続エラー
  - Microsoft Access Database Engine (ODBC Driver) の導入状況を確認
