# mdb2CSV.py 仕様書

## 1. 目的
Access MDBファイル（`.mdb`）内の全テーブルをCSVとして出力する。

## 2. 対象ファイル
- メインスクリプト: `mdb2CSV.py`
- 依存定義: `requirements.txt`

## 3. 動作環境・依存ライブラリ
- Python 3.x
- `pyodbc`
- `tkinter`（標準ライブラリ）
- `tkinterdnd2`（任意。利用可能時はドラッグ&ドロップUIを有効化）
- Microsoft Access Database Engine (ODBC Driver)

## 4. 入出力仕様

### 4.1 入力
- MDBファイルパス（GUIで選択またはドラッグ&ドロップ）
- 対応拡張子: `.mdb`

### 4.2 出力
- 全テーブルを1テーブル1CSVで保存
- 保存先: 入力MDBと同じディレクトリ配下の `<MDBファイル名>/`
- 例: `sample.mdb` -> `sample/T_マスタ.csv`
- CSV仕様: UTF-8 BOM、ヘッダ行あり
- （任意）実行レポートを `.json` として保存
- JSON保存先: 入力MDBと同じディレクトリ配下の `<MDBファイル名>_report.json`
- JSON記録項目: `timestamp`, `target_file`, `status`, `exported_count`, `exported_files`, `output_dir`, `tables_sorted_by_first_column`, `tables_without_sort_key`, `warning_messages`, `message`

## 5. 処理フロー
1. GUI起動時に `tkinterdnd2` の利用可否を判定する。
1. 利用可能ならドラッグ&ドロップ対応ウィンドウを表示、不可ならチェックボックス付き起動画面を表示する。
1. ファイルパスを受け取り、存在チェックを行う。
1. MDB接続を作成し、テーブル一覧を取得する。
1. 各テーブルについてソートキーを決定する（主キー -> unique index -> 先頭列 -> なし）。
1. `SELECT` を実行し、CSVを書き出す。
1. 実行レポート出力チェックボックスがONの場合、`<MDBファイル名>_report.json` に実行結果を追記する。
1. 成功・失敗メッセージをポップアップ表示する。

## 6. 並び順仕様（レコード順）
テーブル内のレコード順は次の優先順で決定する。

1. `primaryKeys()` で主キー列を取得できた場合: 主キー列順で `ORDER BY`
1. 主キー未取得の場合: `statistics(unique=True)` の候補を代替キーとして `ORDER BY`
1. 代替キーも未取得の場合: 先頭列で `ORDER BY`
1. 列情報も取得不可の場合: `ORDER BY` なし

注: Access画面の見え方（保存済み並べ替え等）と一致しない場合がある。

## 7. 関数仕様

### 7.1 `sanitize_filename(name, default_name="table", max_length=120)`
- 目的: Windowsで安全なファイル名に正規化する。
- 主な仕様:
  - 禁止文字（`<>:"/\\|?*` と制御文字）を `_` に置換
  - 末尾の空白/ピリオドを除去
  - 空文字の場合は `default_name` を採用
  - 予約名（`CON`, `PRN`, `AUX`, `NUL`, `COM1`-`COM9`, `LPT1`-`LPT9`）を回避
  - 最大長を `max_length`（既定120）に制限

### 7.2 `build_unique_save_path(output_dir, raw_name, used_names)`
- 目的: 出力CSV名の重複を回避する。
- 主な仕様:
  - 同名がある場合は `_1`, `_2`, ... を付与
  - 重複判定は大文字小文字を区別しない
  - 戻り値は `<output_dir>/<name>.csv`

### 7.3 `get_table_names_in_mdb_order(cursor)`
- 目的: MDB内のテーブル順を取得する。
- 主な仕様:
  - `MSysObjects` を `ORDER BY Id` で参照
  - 参照不可時は `cursor.tables(tableType="TABLE")` にフォールバック
  - システムテーブル（`MSys*`, `USys*`, `~*`）は除外

### 7.4 `get_primary_key_columns(cursor, table_name)`
- 目的: 主キー列（または代替キー列）を取得する。
- 主な仕様:
  - `primaryKeys()` を列名ベースで解釈
  - 未取得時に `statistics(unique=True)` から代替キー候補を選定
  - 複合キーは `KEY_SEQ` / `ORDINAL_POSITION` 順を尊重

### 7.5 `get_table_column_names(cursor, table_name)`
- 目的: 先頭列ソート用にテーブル列一覧を取得する。
- 主な仕様:
  - `cursor.columns(table=...)` から列名を収集
  - 取得できた先頭列を最終フォールバックのソートキーに使う

### 7.6 `build_warning_messages(tables_sorted_by_first_column, tables_without_sort_key, max_items=5)`
- 目的: 実行結果メッセージの注意文を生成する。
- 主な仕様:
  - ポップアップ用は件数制限あり（既定5件）
  - レポート用は `max_items=None` で全件を記録可能

### 7.7 `write_export_report(...)`
- 目的: 実行対象MDBと同じディレクトリにJSONレポートを追記する。
- 主な仕様:
  - レポート名は `<MDBファイル名>_report.json`
  - 既存レポートが配列形式なら末尾追記
  - 既存レポートが単一オブジェクトなら配列化して追記
  - 既存レポート破損時は新規配列として再作成
  - `warning_messages` は全件版を保存する

### 7.8 `export_mdb_tables_to_csv(file_path)`
- 目的: MDB出力本処理を実行する。
- 戻り値:
  - `(success, message, exported_count, output_dir, exported_files, tables_sorted_by_first_column, tables_without_sort_key, warning_messages, report_message)`
- 成功時:
  - 出力件数、保存先、出力CSV一覧、注意情報を返す
- 失敗時:
  - エラー内容をメッセージで返す

### 7.9 `main()`
- 目的: GUI制御とユーザー操作を受け付ける。
- 主な仕様:
  - D&D UI利用可なら専用画面を表示
  - D&D不可でも、チェックボックス付きの起動画面を表示
  - 実行レポート出力チェックボックスでJSON出力の有無を切り替える
  - 結果は `messagebox.showinfo/showerror` で通知

## 8. UI仕様

### 8.1 D&D利用可能時
- ウィンドウタイトル: `MDB to CSV`
- 固定サイズ: `520x250`
- 構成:
  - 説明ラベル（ドラッグ&ドロップ案内）
  - ドロップ領域（`Drop Here`）
  - 対応拡張子ヒント
  - 実行レポート出力チェックボックス（JSON）
  - `ファイル選択...` ボタン（手動選択）

### 8.2 フォールバック時
- ウィンドウタイトル: `MDB to CSV`
- 固定サイズ: `420x150`
- 構成:
  - 説明ラベル
  - 実行レポート出力チェックボックス
  - `ファイル選択...` ボタン

## 9. エラー・例外仕様
- `pyodbc` 未導入: 起動時エラーダイアログを表示して終了
- MDB接続失敗: ドライバ未導入の可能性を含めて通知
- ドロップデータ不正: 「ドロップされたパスを取得できませんでした。」
- 非対応拡張子/ファイル不存在: 「対応していないファイル形式、またはファイルが存在しません。」
- テーブルなし: 「出力対象のテーブルが見つかりませんでした。」
- その他例外: 詳細メッセージ付きで通知

## 10. 既知の制約
- Access画面上の表示順とCSV出力順は一致保証しない。
- ODBCメタデータ取得状況により、主キー検出に失敗するテーブルがある。
- 複数ファイル同時ドロップ時は、対応拡張子かつ実在ファイルを順次処理する。

## 11. 実行方法
```bash
pip install -r requirements.txt
python mdb2CSV.py
```
