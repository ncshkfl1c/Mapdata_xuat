from flask import Flask, request, jsonify, send_file
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl.utils import get_column_letter
import base64
import traceback

app = Flask(__name__)

# =========================================================
# CONFIG
# =========================================================
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024


# =========================================================
# BASE64
# =========================================================
def base64_to_file(base64_string):

    if not base64_string:
        return None

    try:

        # remove data:application/...
        if "," in base64_string:
            base64_string = base64_string.split(",")[1]

        file_bytes = base64.b64decode(base64_string)

        return BytesIO(file_bytes)

    except Exception:
        return None


# =========================================================
# CLEAN KEY
# =========================================================
def clean_key(val):

    if pd.isna(val):
        return ""

    s = str(val).strip()

    if s.endswith(".0"):
        s = s[:-2]

    return s


# =========================================================
# DATE
# =========================================================
def parse_date(val):

    if val in [None, "", "nan"]:
        return None

    try:

        # excel serial
        if isinstance(val, (int, float)):

            d = pd.to_datetime(
                val,
                origin="1899-12-30",
                unit="D",
                errors="coerce"
            )

            if pd.isna(d):
                return None

            return d.to_pydatetime()

        # normal date
        d = pd.to_datetime(
            str(val),
            dayfirst=True,
            errors="coerce"
        )

        if pd.isna(d):
            return None

        return d.to_pydatetime()

    except Exception:
        return None


def safe_format(d):

    if not d:
        return ""

    return d.strftime("%d/%m/%Y")


# =========================================================
# HISTORY
# =========================================================
def fix_serial_date(text):

    if not text:
        return ""

    parts = str(text).split(";")

    result = []

    for p in parts:

        val = p.strip()

        if not val:
            continue

        d = parse_date(val)

        if d:
            val = safe_format(d)

        if val not in result:
            result.append(val)

    return "; ".join(result)


def count_gia_han(text):

    if not text:
        return "0"

    return str(
        len(
            [
                x for x in text.split(";")
                if x.strip()
            ]
        )
    )


# =========================================================
# ENSURE COL
# =========================================================
def ensure_columns(df, total_cols):

    if df.shape[1] < total_cols:

        for i in range(df.shape[1], total_cols):
            df[i] = ""

    return df


# =========================================================
# API
# =========================================================
@app.route("/dongbo", methods=["POST"])
def dongbo():

    try:

        # =================================================
        # GET JSON
        # =================================================
        data = request.get_json()

        if not data:
            return jsonify({
                "success": False,
                "error": "Missing JSON body"
            }), 400

        # =================================================
        # GET BASE64
        # =================================================
        file_old_b64 = data.get("file_old")
        file_new_b64 = data.get("file_new")
        file_map_b64 = data.get("file_map")

        if not file_old_b64:
            return jsonify({
                "success": False,
                "error": "Missing file_old"
            }), 400

        if not file_new_b64:
            return jsonify({
                "success": False,
                "error": "Missing file_new"
            }), 400

        # =================================================
        # CONVERT TO MEMORY FILE
        # =================================================
        file_old = base64_to_file(file_old_b64)
        file_new = base64_to_file(file_new_b64)

        if not file_old:
            return jsonify({
                "success": False,
                "error": "Invalid file_old"
            }), 400

        if not file_new:
            return jsonify({
                "success": False,
                "error": "Invalid file_new"
            }), 400

        file_map = None

        if file_map_b64:
            file_map = base64_to_file(file_map_b64)

        # =================================================
        # READ EXCEL
        # =================================================
        df_old = pd.read_excel(file_old).fillna("")
        df_new = pd.read_excel(file_new).fillna("")

        # reset column index
        df_old.columns = range(df_old.shape[1])
        df_new.columns = range(df_new.shape[1])

        if df_old.empty:
            return jsonify({
                "success": False,
                "error": "file_old empty"
            }), 400

        if df_new.empty:
            return jsonify({
                "success": False,
                "error": "file_new empty"
            }), 400

        # =================================================
        # SKIP 9 ROW FILE NEW
        # =================================================
        df_new = df_new.iloc[9:].reset_index(drop=True)

        # =================================================
        # ENSURE COL
        # =================================================
        df_old = ensure_columns(df_old, 28)

        COL_OLD_ID = 9
        COL_NEW_ID = 11

        if df_new.shape[1] <= COL_NEW_ID:

            return jsonify({
                "success": False,
                "error": "file_new missing ID column"
            }), 400

        # =================================================
        # LOAD NEW DICT
        # =================================================
        dict_all = {}
        dict_new_only = {}

        for i in range(len(df_new)):

            key = clean_key(
                df_new.iloc[i, COL_NEW_ID]
            )

            if key and key not in dict_all:

                dict_all[key] = i
                dict_new_only[key] = i

        # =================================================
        # UPDATE EXISTING
        # =================================================
        rows_keep = []

        for i in reversed(range(len(df_old))):

            colID = clean_key(
                df_old.iloc[i, COL_OLD_ID]
            )

            if colID not in dict_all:
                continue

            rNew = dict_all[colID]

            new_row = df_new.iloc[rNew]

            # =============================================
            # BASIC
            # =============================================
            df_old.iloc[i, 1] = new_row.iloc[3]
            df_old.iloc[i, 2] = new_row.iloc[4]
            df_old.iloc[i, 3] = new_row.iloc[5]
            df_old.iloc[i, 4] = new_row.iloc[6]

            # =============================================
            # DATE
            # =============================================
            ngayMuon = parse_date(
                new_row.iloc[14]
            )

            ngayGiaHan = parse_date(
                new_row.iloc[20]
            )

            if ngayMuon:
                df_old.iloc[i, 12] = (
                    "'" + safe_format(ngayMuon)
                )
            else:
                df_old.iloc[i, 12] = str(
                    new_row.iloc[14]
                )

            df_old.iloc[i, 14] = new_row.iloc[15]

            ngayTra = parse_date(
                new_row.iloc[19]
            )

            if ngayTra:
                df_old.iloc[i, 18] = safe_format(
                    ngayTra
                )
            else:
                df_old.iloc[i, 18] = ""

            # =============================================
            # HISTORY
            # =============================================
            lichSuCu = str(
                df_old.iloc[i, 20]
            ).replace("'", "")

            lichSuCu = fix_serial_date(
                lichSuCu
            )

            arr = [
                x.strip()
                for x in lichSuCu.split(";")
                if x.strip()
            ]

            if (
                ngayGiaHan and
                ngayMuon and
                safe_format(ngayGiaHan)
                != safe_format(ngayMuon)
            ):

                val = safe_format(
                    ngayGiaHan
                )

                if val not in arr:
                    arr.append(val)

            lichSuCu = "; ".join(arr)

            if lichSuCu:
                df_old.iloc[i, 20] = (
                    "'" + lichSuCu
                )
            else:
                df_old.iloc[i, 20] = ""

            df_old.iloc[i, 19] = count_gia_han(
                lichSuCu
            )

            dict_new_only.pop(
                colID,
                None
            )

            rows_keep.append(i)

        # =================================================
        # KEEP MATCHED ROWS
        # =================================================
        if rows_keep:

            df_old = df_old.iloc[
                sorted(rows_keep)
            ].reset_index(drop=True)

        # =================================================
        # ADD NEW ROWS
        # =================================================
        new_rows = []

        for colID, rNew in dict_new_only.items():

            row = [""] * 28

            new_row = df_new.iloc[rNew]

            row[1] = new_row.iloc[3]
            row[2] = new_row.iloc[4]
            row[3] = new_row.iloc[5]
            row[4] = new_row.iloc[6]
            row[9] = new_row.iloc[11]

            # =============================================
            # DATE
            # =============================================
            ngayMuon = parse_date(
                new_row.iloc[14]
            )

            if ngayMuon:
                row[12] = (
                    "'" + safe_format(ngayMuon)
                )

            row[14] = new_row.iloc[15]

            ngayTra = parse_date(
                new_row.iloc[19]
            )

            if ngayTra:
                row[18] = safe_format(
                    ngayTra
                )

            ngayGiaHan = parse_date(
                new_row.iloc[20]
            )

            if (
                ngayGiaHan and
                ngayMuon and
                safe_format(ngayGiaHan)
                != safe_format(ngayMuon)
            ):

                row[20] = (
                    "'" + safe_format(
                        ngayGiaHan
                    )
                )

            # =============================================
            # STT
            # =============================================
            valX = str(
                new_row.iloc[23]
            )

            if valX.isdigit():
                row[24] = (
                    "'" + valX.zfill(10)
                )
            else:
                row[24] = valX

            row[19] = count_gia_han(
                str(row[20]).replace("'", "")
            )

            new_rows.append(row)

        # =================================================
        # CONCAT
        # =================================================
        if new_rows:

            df_add = pd.DataFrame(
                new_rows,
                columns=df_old.columns
            )

            df_old = pd.concat(
                [df_old, df_add],
                ignore_index=True
            )

        # =================================================
        # MAP FILE
        # =================================================
        if file_map:

            df_map = pd.read_excel(
                file_map
            ).fillna("")

            df_map.columns = range(
                df_map.shape[1]
            )

            dict_map = {}

            for i in range(len(df_map)):

                key = clean_key(
                    df_map.iloc[i, 2]
                )

                if (
                    key and
                    key not in dict_map
                ):

                    dict_map[key] = i

            for i in range(len(df_old)):

                key = clean_key(
                    df_old.iloc[i, 1]
                )

                if key not in dict_map:
                    continue

                rMap = dict_map[key]

                df_old.iloc[i, 26] = (
                    df_map.iloc[rMap, 3]
                )

                df_old.iloc[i, 25] = (
                    df_map.iloc[rMap, 4]
                )

        # =================================================
        # OVERDUE
        # =================================================
        today = datetime.today().date()

        for i in range(len(df_old)):

            d = parse_date(
                df_old.iloc[i, 18]
            )

            if d:
                d = d.date()

            df_old.iloc[i, 27] = (
                "Qua han"
                if d and d < today
                else "Chua qua han"
            )

        # =================================================
        # HEADER
        # =================================================
        cols = list(df_old.columns)

        if len(cols) > 27:
            cols[27] = "Tinh trang"

        df_old.columns = cols

        # =================================================
        # EXPORT
        # =================================================
        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            df_old.to_excel(
                writer,
                index=False
            )

            ws = writer.book.active

            # =============================================
            # HIDE COLUMN
            # =============================================
            cols_hide = [
                "A", "G", "H", "I",
                "K", "L", "N", "P",
                "Q", "R", "V", "W", "X"
            ]

            for col in cols_hide:
                ws.column_dimensions[
                    col
                ].hidden = True

            # =============================================
            # AUTO WIDTH
            # =============================================
            for col_idx, col_cells in enumerate(
                ws.columns,
                1
            ):

                max_length = 0

                for cell in col_cells:

                    try:

                        if cell.value:

                            max_length = max(
                                max_length,
                                len(str(cell.value))
                            )

                    except Exception:
                        pass

                adjusted_width = min(
                    max_length + 2,
                    40
                )

                ws.column_dimensions[
                    get_column_letter(col_idx)
                ].width = adjusted_width

        output.seek(0)

        # =================================================
        # RETURN EXCEL
        # =================================================
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="result.xlsx"
        )

    except Exception as e:

        traceback.print_exc()

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


# =========================================================
# HEALTH CHECK
# =========================================================
@app.route("/", methods=["GET"])
def home():

    return jsonify({
        "success": True,
        "message": "API running"
    })


# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5001
    )
