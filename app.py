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

        # remove data:...
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

        s = str(val).strip()

        # dd/mm/yyyy
        if "/" in s:

            arr = s.split("/")

            if len(arr) == 3:

                d = datetime(
                    int(arr[2]),
                    int(arr[1]),
                    int(arr[0])
                )

                return d

        # excel serial
        if s.replace(".", "").isdigit():

            num = float(s)

            if num > 30000:

                d = pd.to_datetime(
                    num,
                    origin="1899-12-30",
                    unit="D",
                    errors="coerce"
                )

                if not pd.isna(d):
                    return d.to_pydatetime()

        # fallback
        d = pd.to_datetime(
            s,
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
def count_history(text):

    if not text:
        return ""

    arr = str(text).split(";")

    count = 0

    for x in arr:

        if str(x).strip():
            count += 1

    return str(count)


# =========================================================
# ENSURE COL
# =========================================================
def ensure_columns(df, total_cols):

    if df.shape[1] < total_cols:

        for i in range(df.shape[1], total_cols):
            df[i] = ""

    return df


# =========================================================
# HOME
# =========================================================
@app.route("/", methods=["GET"])
def home():

    return jsonify({
        "success": True,
        "message": "API running"
    })


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
        # CONVERT FILE
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
        df_old = pd.read_excel(
            file_old,
            dtype=str
        ).fillna("")

        df_new = pd.read_excel(
            file_new,
            dtype=str
        ).fillna("")

        # reset column
        df_old.columns = range(df_old.shape[1])
        df_new.columns = range(df_new.shape[1])

        # =================================================
        # SKIP 9 ROW FILE NEW
        # VBA DATA START ROW 10
        # =================================================
        df_new = df_new.iloc[9:].reset_index(drop=True)

        # =================================================
        # ENSURE COL
        # =================================================
        df_old = ensure_columns(df_old, 28)

        # =================================================
        # COLUMN MAP
        # =================================================
        COL_OLD_ID = 9     # J
        COL_NEW_ID = 11    # L

        # =================================================
        # LOAD NEW DICT
        # =================================================
        dict_all = {}
        dict_new_only = {}

        for i in range(len(df_new)):

            key = clean_key(
                df_new.iloc[i, COL_NEW_ID]
            )

            if key:

                dict_all[key] = i
                dict_new_only[key] = i

        # =================================================
        # UPDATE EXISTING
        # VBA:
        # IF NOT EXISTS -> DELETE
        # =================================================
        rows_keep = []

        for i in range(len(df_old)):

            colID = clean_key(
                df_old.iloc[i, COL_OLD_ID]
            )

            # NOT EXISTS -> DELETE
            if colID not in dict_all:
                continue

            rNew = dict_all[colID]

            new_row = df_new.iloc[rNew]

            # =================================================
            # BASIC
            # =================================================
            df_old.iloc[i, 1] = new_row.iloc[3]   # B <- D
            df_old.iloc[i, 2] = new_row.iloc[4]   # C <- E
            df_old.iloc[i, 3] = new_row.iloc[5]   # D <- F
            df_old.iloc[i, 4] = new_row.iloc[6]   # E <- G

            # =================================================
            # M <- O
            # =================================================
            tmp = str(
                new_row.iloc[14]
            ).strip()

            ngayMuon = ""

            if tmp:

                d = parse_date(tmp)

                if d:

                    df_old.iloc[i, 12] = (
                        "'" + safe_format(d)
                    )

                    ngayMuon = safe_format(d)

                else:

                    df_old.iloc[i, 12] = tmp
                    ngayMuon = tmp

            # =================================================
            # O <- P
            # =================================================
            df_old.iloc[i, 14] = (
                new_row.iloc[15]
            )

            # =================================================
            # S <- T
            # =================================================
            tmp = str(
                new_row.iloc[19]
            ).strip()

            if tmp:

                d = parse_date(tmp)

                if d:
                    df_old.iloc[i, 18] = safe_format(d)
                else:
                    df_old.iloc[i, 18] = tmp

            else:

                df_old.iloc[i, 18] = ""

            # =================================================
            # U <- U
            # =================================================
            tmp = str(
                new_row.iloc[20]
            ).strip()

            ngayGiaHanNew = ""

            if tmp:

                d = parse_date(tmp)

                if d:
                    ngayGiaHanNew = safe_format(d)
                else:
                    ngayGiaHanNew = tmp

            # =================================================
            # HISTORY
            # =================================================
            lichSuCu = str(
                df_old.iloc[i, 20]
            )

            if lichSuCu.startswith("'"):
                lichSuCu = lichSuCu[1:]

            if (
                ngayGiaHanNew and
                ngayGiaHanNew != ngayMuon
            ):

                if ngayGiaHanNew not in lichSuCu:

                    if lichSuCu == "":

                        df_old.iloc[i, 20] = (
                            "'" + ngayGiaHanNew
                        )

                    else:

                        df_old.iloc[i, 20] = (
                            "'" +
                            lichSuCu +
                            "; " +
                            ngayGiaHanNew
                        )

            # =================================================
            # COUNT T
            # =================================================
            lichSuCu = str(
                df_old.iloc[i, 20]
            )

            if lichSuCu.startswith("'"):
                lichSuCu = lichSuCu[1:]

            df_old.iloc[i, 19] = (
                count_history(lichSuCu)
            )

            # remove new only
            dict_new_only.pop(
                colID,
                None
            )

            # keep row
            rows_keep.append(i)

        # =================================================
        # DELETE ROW NOT EXISTS
        # SAME VBA DELETE
        # =================================================
        df_old = df_old.iloc[
            rows_keep
        ].reset_index(drop=True)

        # =================================================
        # ADD NEW
        # =================================================
        new_rows = []

        for colID, rNew in dict_new_only.items():

            row = [""] * 28

            new_row = df_new.iloc[rNew]

            # =================================================
            # BASIC
            # =================================================
            row[1] = new_row.iloc[3]
            row[2] = new_row.iloc[4]
            row[3] = new_row.iloc[5]
            row[4] = new_row.iloc[6]
            row[9] = new_row.iloc[11]

            # =================================================
            # M <- O
            # =================================================
            tmp = str(
                new_row.iloc[14]
            ).strip()

            ngayMuon = ""

            if tmp:

                d = parse_date(tmp)

                if d:

                    row[12] = (
                        "'" + safe_format(d)
                    )

                    ngayMuon = safe_format(d)

                else:

                    row[12] = tmp
                    ngayMuon = tmp

            # =================================================
            # O <- P
            # =================================================
            row[14] = new_row.iloc[15]

            # =================================================
            # S <- T
            # =================================================
            tmp = str(
                new_row.iloc[19]
            ).strip()

            if tmp:

                d = parse_date(tmp)

                if d:
                    row[18] = safe_format(d)
                else:
                    row[18] = tmp

            # =================================================
            # U <- U
            # =================================================
            tmp = str(
                new_row.iloc[20]
            ).strip()

            if tmp:

                d = parse_date(tmp)

                if d:

                    ngayGiaHanNew = safe_format(d)

                else:

                    ngayGiaHanNew = tmp

                if ngayGiaHanNew != ngayMuon:

                    row[20] = (
                        "'" + ngayGiaHanNew
                    )

            # =================================================
            # COUNT T
            # =================================================
            lichSuCu = str(row[20])

            if lichSuCu.startswith("'"):
                lichSuCu = lichSuCu[1:]

            row[19] = count_history(
                lichSuCu
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
                file_map,
                dtype=str
            ).fillna("")

            df_map.columns = range(
                df_map.shape[1]
            )

            dict_map = {}

            for i in range(len(df_map)):

                keyMap = clean_key(
                    df_map.iloc[i, 2]
                )

                if keyMap:
                    dict_map[keyMap] = i

            for i in range(len(df_old)):

                keyMap = clean_key(
                    df_old.iloc[i, 1]
                )

                if keyMap in dict_map:

                    rMap = dict_map[keyMap]

                    # Z <- E
                    df_old.iloc[i, 25] = (
                        df_map.iloc[rMap, 4]
                    )

                    # AA <- D
                    df_old.iloc[i, 26] = (
                        df_map.iloc[rMap, 3]
                    )

        # =================================================
        # QUA HAN
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

            # =================================================
            # HIDE COLUMN
            # =================================================
            cols_hide = [
                "A", "G", "H", "I",
                "K", "L", "N", "P",
                "Q", "R", "V", "W", "X"
            ]

            for col in cols_hide:

                ws.column_dimensions[
                    col
                ].hidden = True

            # =================================================
            # AUTO WIDTH
            # =================================================
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

                    except:
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
        # RETURN FILE
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
# MAIN
# =========================================================
if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5001
    )
