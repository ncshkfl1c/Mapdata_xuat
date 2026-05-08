from flask import Flask, request, jsonify
import pandas as pd
from datetime import datetime
import tempfile
from openpyxl import load_workbook
from io import BytesIO
import base64
import traceback

app = Flask(__name__)

# =========================================================
# CONFIG
# =========================================================
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024

# =========================================================
# BASE64
# =========================================================
def b64_to_file(b64):

    if not b64:
        return None

    try:

        if "," in b64:
            b64 = b64.split(",")[1]

        return BytesIO(
            base64.b64decode(b64)
        )

    except:
        return None


# =========================================================
# CLEAN KEY
# =========================================================
def clean_key(val):

    if val is None:
        return ""

    s = str(val).strip()

    if s.endswith(".0"):
        s = s[:-2]

    try:

        if "E" in s or "e" in s:
            s = str(int(float(s)))

    except:
        pass

    return s


# =========================================================
# DATE
# =========================================================
def parse_date(val):

    if val is None:
        return None

    s = str(val).strip()

    if s == "":
        return None

    # =====================================================
    # dd/mm/yyyy
    # =====================================================
    try:

        return datetime.strptime(
            s,
            "%d/%m/%Y"
        )

    except:
        pass

    # =====================================================
    # excel serial
    # =====================================================
    try:

        d = pd.to_datetime(
            val,
            origin='1899-12-30',
            unit='D',
            errors='coerce'
        )

        if pd.isna(d):
            return None

        return d.to_pydatetime()

    except:
        pass

    # =====================================================
    # fallback
    # =====================================================
    try:

        d = pd.to_datetime(
            val,
            errors="coerce"
        )

        if pd.isna(d):
            return None

        return d.to_pydatetime()

    except:

        return None


def safe_format(d):

    return d.strftime(
        "%d/%m/%Y"
    ) if d else ""


# =========================================================
# HISTORY
# =========================================================
def fix_serial_date(text):

    if not text:
        return ""

    parts = str(text).split(";")

    result = []

    for p in parts:

        d = parse_date(p.strip())

        val = (
            d.strftime("%d/%m/%Y")
            if d
            else p.strip()
        )

        if val and val not in result:
            result.append(val)

    return "; ".join(result)


def count_gia_han(text):

    if not text:
        return ""

    return str(
        len([
            x for x in text.split(";")
            if x.strip()
        ])
    )


# =========================================================
# ENSURE COL
# =========================================================
def ensure_columns(df, total_cols):

    if df.shape[1] < total_cols:

        for i in range(
            df.shape[1],
            total_cols
        ):

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
        file_old = b64_to_file(
            file_old_b64
        )

        file_new = b64_to_file(
            file_new_b64
        )

        file_map = b64_to_file(
            file_map_b64
        )

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

        # =================================================
        # SKIP 9 ROW FILE NEW
        # =================================================
        df_new = df_new.iloc[9:].reset_index(drop=True)

        # =================================================
        # ENSURE COL
        # =================================================
        df_old = ensure_columns(
            df_old,
            28
        )

        COL_OLD_ID = 9
        COL_NEW_ID = 11

        dict_all = {}
        dict_new_only = {}

        # =================================================
        # LOAD NEW
        # =================================================
        for i in range(len(df_new)):

            key = clean_key(
                df_new.iloc[i, COL_NEW_ID]
            )

            if key:

                dict_all[key] = i
                dict_new_only[key] = i

        # =================================================
        # KEEP HEADER
        # =================================================
        rows_keep = [0]

        # =================================================
        # UPDATE
        # =================================================
        for i in reversed(
            range(1, len(df_old))
        ):

            colID = clean_key(
                df_old.iloc[i, COL_OLD_ID]
            )

            if colID in dict_all:

                rNew = dict_all[colID]

                # =========================================
                # BASIC
                # =========================================
                df_old.iloc[i, 1] = df_new.iloc[rNew, 3]
                df_old.iloc[i, 2] = df_new.iloc[rNew, 4]
                df_old.iloc[i, 3] = df_new.iloc[rNew, 5]
                df_old.iloc[i, 4] = df_new.iloc[rNew, 6]

                # =========================================
                # DATE
                # =========================================
                ngayMuon = parse_date(
                    df_new.iloc[rNew, 14]
                )

                ngayGiaHan = parse_date(
                    df_new.iloc[rNew, 20]
                )

                df_old.iloc[i, 12] = (
                    "'" + safe_format(ngayMuon)
                    if ngayMuon
                    else df_new.iloc[rNew, 14]
                )

                df_old.iloc[i, 14] = (
                    df_new.iloc[rNew, 15]
                )

                df_old.iloc[i, 18] = safe_format(
                    parse_date(
                        df_new.iloc[rNew, 19]
                    )
                )

                # =========================================
                # HISTORY
                # =========================================
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
                    ngayGiaHan != ngayMuon
                ):

                    val = safe_format(
                        ngayGiaHan
                    )

                    if val not in arr:
                        arr.append(val)

                lichSuCu = "; ".join(arr)

                df_old.iloc[i, 20] = (
                    "'" + lichSuCu
                    if lichSuCu
                    else ""
                )

                df_old.iloc[i, 19] = (
                    count_gia_han(
                        lichSuCu
                    )
                )

                # =========================================
                # REMOVE
                # =========================================
                dict_new_only.pop(
                    colID,
                    None
                )

                rows_keep.append(i)

        # =================================================
        # DELETE ROW NOT MATCH
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

            row[1] = df_new.iloc[rNew, 3]
            row[2] = df_new.iloc[rNew, 4]
            row[3] = df_new.iloc[rNew, 5]
            row[4] = df_new.iloc[rNew, 6]
            row[9] = df_new.iloc[rNew, 11]

            ngayMuon = parse_date(
                df_new.iloc[rNew, 14]
            )

            ngayGiaHan = parse_date(
                df_new.iloc[rNew, 20]
            )

            # =============================================
            # M
            # =============================================
            if ngayMuon:

                row[12] = (
                    "'" + safe_format(
                        ngayMuon
                    )
                )

            # =============================================
            # O
            # =============================================
            row[14] = (
                df_new.iloc[rNew, 15]
            )

            # =============================================
            # S
            # =============================================
            ngayTra = parse_date(
                df_new.iloc[rNew, 19]
            )

            if ngayTra:

                row[18] = safe_format(
                    ngayTra
                )

            # =============================================
            # U
            # =============================================
            if (
                ngayGiaHan and
                ngayMuon and
                ngayGiaHan != ngayMuon
            ):

                row[20] = (
                    "'" +
                    safe_format(
                        ngayGiaHan
                    )
                )

            # =============================================
            # X
            # =============================================
            valX = str(
                df_new.iloc[rNew, 23]
            )

            row[24] = (
                "'" + valX.zfill(10)
                if valX.isdigit()
                else valX
            )

            # =============================================
            # T
            # =============================================
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
        # MAP
        # =================================================
        if file_map:

            df_map = pd.read_excel(
                file_map,
                dtype=str
            ).fillna("")

            dict_map = {}

            for i in range(len(df_map)):

                key = clean_key(
                    df_map.iloc[i, 2]
                )

                if key:
                    dict_map[key] = i

            for i in range(
                1,
                len(df_old)
            ):

                key = clean_key(
                    df_old.iloc[i, 1]
                )

                if key in dict_map:

                    rMap = dict_map[key]

                    df_old.iloc[i, 26] = (
                        df_map.iloc[rMap, 3]
                    )

                    df_old.iloc[i, 25] = (
                        df_map.iloc[rMap, 4]
                    )

        # =================================================
        # QUA HAN
        # =================================================
        today = datetime.today()

        for i in range(len(df_old)):

            d = parse_date(
                df_old.iloc[i, 18]
            )

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
            cols[27] = "Tình trạng"

        df_old.columns = cols

        # =================================================
        # EXPORT
        # =================================================
        tmp = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        )

        path = tmp.name

        tmp.close()

        df_old.to_excel(
            path,
            index=False,
            engine="openpyxl"
        )

        # =================================================
        # HIDE COLUMN
        # =================================================
        wb = load_workbook(path)

        ws = wb.active

        cols_hide = [
            "A", "G", "H", "I",
            "K", "L", "N", "P",
            "Q", "R", "V", "W", "X"
        ]

        for col in cols_hide:

            ws.column_dimensions[
                col
            ].hidden = True

        wb.save(path)

        # =================================================
        # RETURN BASE64
        # =================================================
        with open(path, "rb") as f:

            file_data = base64.b64encode(
                f.read()
            ).decode("utf-8")

        return jsonify({
            "success": True,
            "filename": "result.xlsx",
            "file": file_data
        })

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
        debug=True,
        host="0.0.0.0",
        port=5001
    )
