import csv
import json
import os
import re
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox

try:
    import pyodbc
    PYODBC_IMPORT_ERROR = None
except Exception as e:
    pyodbc = None
    PYODBC_IMPORT_ERROR = e

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkinterDnD = None

WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}
SUPPORTED_EXTENSIONS = (".mdb",)
SYSTEM_TABLE_PREFIXES = ("msys", "usys", "~")


def sanitize_filename(name, default_name="table", max_length=120):
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", str(name))
    safe = safe.strip().rstrip(".")

    if not safe:
        safe = default_name

    if safe.split(".")[0].upper() in WINDOWS_RESERVED_NAMES:
        safe = f"_{safe}"

    if len(safe) > max_length:
        safe = safe[:max_length].rstrip(" .")

    return safe or default_name


def build_unique_save_path(output_dir, raw_name, used_names):
    base_name = sanitize_filename(raw_name)
    candidate = base_name
    index = 1

    while candidate.lower() in used_names:
        suffix = f"_{index}"
        allowed = max(1, 120 - len(suffix))
        candidate = f"{base_name[:allowed]}{suffix}"
        candidate = candidate.rstrip(" .")
        index += 1

    used_names.add(candidate.lower())
    return os.path.join(output_dir, f"{candidate}.csv")


def is_supported_mdb_file(file_path):
    return os.path.isfile(file_path) and file_path.lower().endswith(SUPPORTED_EXTENSIONS)


def parse_dnd_file_paths(root, data):
    return [p for p in root.tk.splitlist(data) if p]


def is_user_table_name(name):
    if not name:
        return False
    lower_name = name.lower()
    return not any(lower_name.startswith(prefix) for prefix in SYSTEM_TABLE_PREFIXES)


def dedupe_keep_order(names):
    seen = set()
    result = []
    for name in names:
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(name)
    return result


def get_table_names_in_mdb_order(cursor):
    """
    MDB内部のテーブル定義順で取得を試みる。
    取得できない場合は ODBC の tables() 結果にフォールバックする。
    """
    try:
        rows = cursor.execute(
            """
            SELECT Name
            FROM MSysObjects
            WHERE Type IN (1, 4, 6)
              AND Name NOT LIKE 'MSys*'
              AND Name NOT LIKE 'USys*'
              AND Name NOT LIKE '~*'
            ORDER BY Id
            """
        ).fetchall()
        names = [row[0] for row in rows if row and is_user_table_name(row[0])]
        names = dedupe_keep_order(names)
        if names:
            return names
    except Exception:
        # MSysObjects 参照不可(権限不足など)の場合はフォールバック
        pass

    table_rows = cursor.tables(tableType="TABLE").fetchall()
    names = [row.table_name for row in table_rows if is_user_table_name(row.table_name)]
    return dedupe_keep_order(names)


def quote_identifier(name):
    return f"[{str(name).replace(']', ']]')}]"


def build_column_index(description):
    index = {}
    if not description:
        return index
    for i, col in enumerate(description):
        name = str(col[0]).strip().lower()
        index[name] = i
    return index


def first_existing_key(index_map, candidates):
    for key in candidates:
        if key in index_map:
            return index_map[key]
    return None


def to_int_or_default(value, default):
    try:
        return int(value)
    except Exception:
        return default


def get_primary_key_columns(cursor, table_name):
    """
    テーブルの主キー列を KEY_SEQ 順で返す。
    取得できない/主キーなしの場合は unique index を代替キーとして試す。
    """
    conn = cursor.connection
    rows = []
    desc = None
    try:
        pk_cursor = conn.cursor()
    except Exception:
        pk_cursor = None

    try:
        if pk_cursor is None:
            pk_cursor = conn.cursor()
        rows = pk_cursor.primaryKeys(table=table_name).fetchall()
        desc = pk_cursor.description
    except Exception:
        rows = []
        desc = None

    cols = []
    index_map = build_column_index(desc)
    col_idx = first_existing_key(index_map, ["column_name", "columnname", "column"])
    seq_idx = first_existing_key(index_map, ["key_seq", "keyseq", "ordinal_position", "ordinalposition"])

    for row in rows:
        col_name = None
        key_seq = None

        if col_idx is not None and len(row) > col_idx:
            col_name = row[col_idx]
        if seq_idx is not None and len(row) > seq_idx:
            key_seq = row[seq_idx]

        # ODBC 標準配置へのフォールバック
        if col_name is None and len(row) > 3:
            col_name = row[3]
        if key_seq is None and len(row) > 4:
            key_seq = row[4]

        if col_name:
            seq = to_int_or_default(key_seq, 10**9)
            cols.append((seq, col_name))

    cols.sort(key=lambda x: x[0])
    primary_key_cols = [name for _, name in cols]
    if primary_key_cols:
        return primary_key_cols

    # primaryKeys が取れないドライバ向けのフォールバック:
    # unique index の先頭候補を利用して順序安定化を試みる
    try:
        st_cursor = conn.cursor()
        st_rows = st_cursor.statistics(table=table_name, unique=True).fetchall()
        st_desc = st_cursor.description
    except Exception:
        return []

    st_index = build_column_index(st_desc)
    idx_name_i = first_existing_key(st_index, ["index_name"])
    col_name_i = first_existing_key(st_index, ["column_name", "columnname", "column"])
    ord_pos_i = first_existing_key(st_index, ["ordinal_position", "ordinalposition", "seq_in_index"])
    non_unique_i = first_existing_key(st_index, ["non_unique", "nonunique"])

    grouped = {}
    for row in st_rows:
        index_name = row[idx_name_i] if idx_name_i is not None and len(row) > idx_name_i else None
        col_name = row[col_name_i] if col_name_i is not None and len(row) > col_name_i else None
        ord_pos = row[ord_pos_i] if ord_pos_i is not None and len(row) > ord_pos_i else None
        non_unique = row[non_unique_i] if non_unique_i is not None and len(row) > non_unique_i else None

        if not index_name or not col_name:
            continue
        if non_unique not in (0, False, "0", None):
            continue

        grouped.setdefault(str(index_name), []).append((to_int_or_default(ord_pos, 10**9), col_name))

    if not grouped:
        return []

    def rank(item):
        name, cols_in_index = item
        lname = name.lower()
        primary_hint = 0 if ("primary" in lname or lname == "pk") else 1
        return (primary_hint, len(cols_in_index), lname)

    best_name, best_cols = sorted(grouped.items(), key=rank)[0]
    best_cols.sort(key=lambda x: x[0])
    return [name for _, name in best_cols]


def get_table_column_names(cursor, table_name):
    try:
        rows = cursor.columns(table=table_name).fetchall()
    except Exception:
        return []

    names = []
    for row in rows:
        col_name = getattr(row, "column_name", None)
        if col_name is None and len(row) > 3:
            col_name = row[3]
        if col_name:
            names.append(col_name)
    return dedupe_keep_order(names)


def build_select_query(table_name, order_columns):
    table_expr = quote_identifier(table_name)
    if not order_columns:
        return f"SELECT * FROM {table_expr}"

    order_expr = ", ".join(quote_identifier(col) for col in order_columns)
    return f"SELECT * FROM {table_expr} ORDER BY {order_expr}"


def get_access_connection(file_path):
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={file_path};"
    )
    return pyodbc.connect(conn_str)


def build_warning_messages(tables_sorted_by_first_column, tables_without_sort_key, max_items=5):
    warnings = []
    if tables_sorted_by_first_column:
        items = tables_sorted_by_first_column
        if max_items is not None:
            items = tables_sorted_by_first_column[:max_items]
        limited = ", ".join(items)
        suffix = ""
        if max_items is not None and len(tables_sorted_by_first_column) > max_items:
            suffix = " ..."
        warnings.append(
            "注意: 一部テーブルで主キーを検出できず、先頭列でソートして出力しました。"
            f"\n対象: {limited}{suffix}"
        )

    if tables_without_sort_key:
        items = tables_without_sort_key
        if max_items is not None:
            items = tables_without_sort_key[:max_items]
        limited = ", ".join(items)
        suffix = ""
        if max_items is not None and len(tables_without_sort_key) > max_items:
            suffix = " ..."
        warnings.append(
            "注意: 一部テーブルはソートキーを取得できず、ORDER BY なしで出力しました。"
            f"\n対象: {limited}{suffix}"
        )

    return warnings


def write_export_report(
    file_path,
    success,
    exported_count,
    output_dir,
    message,
    exported_files,
    tables_sorted_by_first_column,
    tables_without_sort_key,
    warning_messages,
):
    dir_path = os.path.dirname(file_path)
    source_name = os.path.splitext(os.path.basename(file_path))[0]
    report_path = os.path.join(dir_path, f"{source_name}_report.json")

    record = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "target_file": file_path,
        "status": "SUCCESS" if success else "FAILED",
        "exported_count": exported_count,
        "exported_files": exported_files,
        "output_dir": output_dir,
        "tables_sorted_by_first_column": tables_sorted_by_first_column,
        "tables_without_sort_key": tables_without_sort_key,
        "warning_messages": warning_messages,
        "message": message.replace("\n", " "),
    }

    logs = []
    if os.path.exists(report_path):
        try:
            with open(report_path, "r", encoding="utf-8-sig") as f:
                loaded = json.load(f)
            if isinstance(loaded, list):
                logs = loaded
            elif isinstance(loaded, dict):
                logs = [loaded]
        except Exception:
            logs = []

    logs.append(record)

    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(logs, f, ensure_ascii=False, indent=2)

    return report_path


def export_mdb_tables_to_csv(file_path):
    if pyodbc is None:
        return (
            False,
            "必要なPythonライブラリ 'pyodbc' が見つかりません。\n"
            "次を実行してインストールしてください:\n"
            "pip install -r requirements.txt\n\n"
            f"詳細: {PYODBC_IMPORT_ERROR}",
            0,
            "",
            [],
            [],
            [],
            [],
            "",
        )

    if not os.path.exists(file_path):
        message = f"ファイルが見つかりません: {file_path}"
        return False, message, 0, "", [], [], [], [], message

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    dir_path = os.path.dirname(file_path)
    output_dir = os.path.join(dir_path, base_name)

    try:
        conn = get_access_connection(file_path)
    except Exception as e:
        return (
            False,
            "MDBへの接続に失敗しました。\n"
            "Microsoft Access Database Engine (ODBC Driver) が未導入の可能性があります。\n"
            f"詳細: {e}",
            0,
            output_dir,
            [],
            [],
            [],
            [],
            "",
        )

    exported_count = 0
    used_names = set()
    exported_files = []
    tables_sorted_by_first_column = []
    tables_without_sort_key = []

    try:
        cursor = conn.cursor()
        table_names = get_table_names_in_mdb_order(cursor)

        if not table_names:
            message = "出力対象のテーブルが見つかりませんでした。"
            return False, message, 0, output_dir, [], [], [], [], message

        os.makedirs(output_dir, exist_ok=True)

        for table_name in table_names:
            save_path = build_unique_save_path(output_dir, table_name, used_names)
            pk_columns = get_primary_key_columns(cursor, table_name)
            order_columns = pk_columns

            if not order_columns:
                col_names = get_table_column_names(cursor, table_name)
                if col_names:
                    order_columns = [col_names[0]]
                    tables_sorted_by_first_column.append(table_name)
                else:
                    tables_without_sort_key.append(table_name)

            query = build_select_query(table_name, order_columns)
            cursor.execute(query)

            columns = [desc[0] for desc in cursor.description] if cursor.description else []
            rows = cursor.fetchall()

            with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                if columns:
                    writer.writerow(columns)
                for row in rows:
                    writer.writerow([value if value is not None else "" for value in row])

            exported_files.append(os.path.basename(save_path))
            exported_count += 1

        base_message = f"{exported_count} テーブルをCSV出力しました。\n保存先: {output_dir}"
        popup_warning_messages = build_warning_messages(
            tables_sorted_by_first_column=tables_sorted_by_first_column,
            tables_without_sort_key=tables_without_sort_key,
            max_items=5,
        )
        warning_messages = build_warning_messages(
            tables_sorted_by_first_column=tables_sorted_by_first_column,
            tables_without_sort_key=tables_without_sort_key,
            max_items=None,
        )

        message = base_message
        for warning_text in popup_warning_messages:
            message += f"\n\n{warning_text}"

        report_message = base_message
        for warning_text in warning_messages:
            report_message += f"\n\n{warning_text}"

        return (
            True,
            message,
            exported_count,
            output_dir,
            exported_files,
            tables_sorted_by_first_column,
            tables_without_sort_key,
            warning_messages,
            report_message,
        )
    except Exception as e:
        if pyodbc is not None and isinstance(e, pyodbc.Error):
            message = f"テーブル出力中にODBCエラーが発生しました。\n詳細: {e}"
            return (
                False,
                message,
                exported_count,
                output_dir,
                exported_files,
                tables_sorted_by_first_column,
                tables_without_sort_key,
                [],
                message,
            )
        message = f"出力中にエラーが発生しました。\n詳細: {e}"
        return (
            False,
            message,
            exported_count,
            output_dir,
            exported_files,
            tables_sorted_by_first_column,
            tables_without_sort_key,
            [],
            message,
        )
    finally:
        conn.close()


def main():
    if pyodbc is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "起動エラー",
            "必要なPythonライブラリ 'pyodbc' が見つかりません。\n"
            "次を実行してインストールしてください:\n"
            "pip install -r requirements.txt",
        )
        root.destroy()
        return

    root = None
    report_enabled = None

    def run_export(file_path):
        (
            success,
            message,
            exported_count,
            output_dir,
            exported_files,
            tables_sorted_by_first_column,
            tables_without_sort_key,
            warning_messages,
            report_message,
        ) = export_mdb_tables_to_csv(file_path)
        report_suffix = ""

        if report_enabled.get():
            report_path = write_export_report(
                file_path=file_path,
                success=success,
                exported_count=exported_count,
                output_dir=output_dir,
                message=report_message,
                exported_files=exported_files,
                tables_sorted_by_first_column=tables_sorted_by_first_column,
                tables_without_sort_key=tables_without_sort_key,
                warning_messages=warning_messages,
            )
            report_suffix = f"\n\nレポート: {report_path}"

        if success:
            messagebox.showinfo("完了", f"{message}{report_suffix}")
        else:
            messagebox.showerror("結果", f"{message}{report_suffix}")

    def browse_file():
        file_path = filedialog.askopenfilename(
            title="CSV出力したいMDBファイルを選択してください",
            filetypes=[("Access MDB Files", "*.mdb"), ("All Files", "*.*")],
        )
        if file_path:
            run_export(file_path)

    if TkinterDnD is not None and DND_FILES is not None:
        root = TkinterDnD.Tk()
        report_enabled = tk.BooleanVar(master=root, value=False)

        root.title("MDB to CSV")
        root.geometry("520x250")
        root.resizable(False, False)

        title = tk.Label(root, text="MDBファイルをここにドラッグ&ドロップ")
        title.pack(pady=(16, 8))

        drop_area = tk.Label(
            root,
            text="Drop Here",
            relief="groove",
            bd=2,
            width=52,
            height=5,
        )
        drop_area.pack(padx=16, pady=8, fill="x")

        hint = tk.Label(root, text="対応形式: .mdb")
        hint.pack(pady=(4, 8))

        report_checkbox = tk.Checkbutton(
            root,
            text="実行レポートを出力する（同じフォルダにJSON）",
            variable=report_enabled,
        )
        report_checkbox.pack(pady=(0, 8))

        browse_btn = tk.Button(root, text="ファイル選択...", command=browse_file)
        browse_btn.pack(pady=(0, 12))

        def on_drop(event):
            paths = parse_dnd_file_paths(root, event.data)
            if not paths:
                messagebox.showerror("結果", "ドロップされたパスを取得できませんでした。")
                return

            target_paths = [p for p in paths if is_supported_mdb_file(p)]
            if not target_paths:
                messagebox.showerror("結果", "対応していないファイル形式、またはファイルが存在しません。")
                return

            for file_path in target_paths:
                run_export(file_path)

        drop_area.drop_target_register(DND_FILES)
        drop_area.dnd_bind("<<Drop>>", on_drop)
        root.mainloop()
        return

    root = tk.Tk()
    report_enabled = tk.BooleanVar(master=root, value=False)

    root.title("MDB to CSV")
    root.geometry("420x150")
    root.resizable(False, False)

    title = tk.Label(root, text="出力対象MDBファイルを選択してください")
    title.pack(pady=(12, 8))

    report_checkbox = tk.Checkbutton(
        root,
        text="実行レポートを出力する（同じフォルダにJSON）",
        variable=report_enabled,
    )
    report_checkbox.pack(pady=(0, 8))

    browse_btn = tk.Button(root, text="ファイル選択...", command=browse_file)
    browse_btn.pack(pady=(0, 12))

    root.mainloop()


if __name__ == "__main__":
    main()
