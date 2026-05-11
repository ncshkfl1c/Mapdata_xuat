from flask import Flask, request, jsonify
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import tempfile
import base64
import os

app = Flask(__name__)

# =========================================================
# CLEAN KEY
# =========================================================
def clean_key(val):

    if val is None:
        return ""

    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")

    s = str(val).strip()

    if s == "":
        return ""

    # remove .0
    if s.endswith(".0"):
        s = s[:-2]

    # scientific notation
    try:
        if "E" in s or "e" in s:
            s = str(int(float(s)))
    except:
        pass

    return s


# =========================================================
# PARSE DATE
# =========================================================
def parse_date(val):

    if val is None:
        return None

    if str(val).strip() == "":
        return None

    # datetime
    if isinstance(val, datetime):
        return val

    # excel serial
    if isinstance(val, (int, float)):

        try:
            d = pd.to_datetime(
                val,
                origin="1899-12-30",
                unit="D",
                errors="coerce"
            )

            if pd.isna(d):
                return None

            return d.to_pydatetime()

        except:
            pass

    s = str(val).strip()

    # dd/mm/yyyy
    try:
        return datetime.strptime(s, "%d/%m/%Y")
    except:
        pass

    # auto
    try:

        d = pd.to_datetime(
            s,
            dayfirst=True,
            errors="coerce"
        )

        if pd.isna(d):
            return None

        return d.to_pydatetime()

    except:
        return None


# =========================================================
# SAFE FORMAT
# =========================================================
def safe_format(d):

    if not d:
        return ""

    return d.strftime("%d/%m/%Y")


# =========================================================
# FIX SERIAL DATE
# =========================================================
def fix_serial_date(text):

    if not text:
        return ""

    result = []

    parts = str(text).split(";")

    for p in parts:

        p = p.strip()

        if not p:
            continue

        d = parse_date(p)

        val = d.strftime("%d/%m/%Y") if d else p

        if val not in result:
            result.append(val)

    return "; ".join(result)


# =========================================================
# COUNT GIA HAN
# =========================================================
def count_gia_han(text):

    if not text:
        return "0"

    return str(
        len([
            x for x in text.split(";")
            if x.strip()
        ])
    )


# =========================================================
# ENSURE COLUMNS
# =========================================================
def ensure_columns(df, total_cols):

    if df.shape[1] < total_cols:

        for i in range(df.shape[1], total_cols):
            df[i] = ""

    return df


# =========================================================
# SAVE BASE64 FILE
# =========================================================
def save_base64_file(base64_string):

    if "," in base64_string:
        base64_string = base64_string.split(",")[1]

    file_data = base64.b64decode(base64_string)

    tmp = tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".xlsx"
    )

    tmp.write(file_data)
    tmp.close()

    return tmp.name


# =========================================================
# MAIN API
# =========================================================
@app.route("/dongbo", methods=["POST"])
def dongbo():

    temp_files = []

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
        # SAVE TEMP FILES
        # =================================================
        path_old = save_base64_file(file_old_b64)
        path_new = save_base64_file(file_new_b64)

        temp_files.extend([path_old, path_new])

        path_map = None

        if file_map_b64:
            path_map = save_base64_file(file_map_b64)
            temp_files.append(path_map)

        # =================================================
        # READ EXCEL
        # =================================================
        # KHÔNG dùng dtype=str
        df_old = pd.read_excel(path_old).fillna("")
        df_new = pd.read_excel(path_new).fillna("")

        # force numeric columns
        df_old.columns = range(df_old.shape[1])
        df_new.columns = range(df_new.shape[1])

        # skip top rows
        df_new = df_new.iloc[9:].reset_index(drop=True)

        # ensure enough columns
        df_old = ensure_columns(df_old, 28)

        # =================================================
        # CONFIG
        # =================================================
        COL_OLD_ID = 9
        COL_NEW_ID = 11

        # =================================================
        # CREATE NEW DICT
        # =================================================
        dict_new = {}

        for i in range(len(df_new)):

            key = clean_key(df_new.iloc[i, COL_NEW_ID])

            if key:
                dict_new[key] = i

        # =================================================
        # UPDATE OLD
        # =================================================
        for i in range(len(df_old)):

            colID = clean_key(df_old.iloc[i, COL_OLD_ID])

            if colID in dict_new:

                rNew = dict_new[colID]

                # =========================================
                # COPY BASIC
                # =========================================
                df_old.iloc[i, 1] = df_new.iloc[rNew, 3]
                df_old.iloc[i, 2] = df_new.iloc[rNew, 4]
                df_old.iloc[i, 3] = df_new.iloc[rNew, 5]
                df_old.iloc[i, 4] = df_new.iloc[rNew, 6]

                # =========================================
                # DATES
                # =========================================
                ngayMuon = parse_date(
                    df_new.iloc[rNew, 14]
                )

                ngayGiaHan = parse_date(
                    df_new.iloc[rNew, 20]
                )

                ngayTra = parse_date(
                    df_new.iloc[rNew, 19]
                )

                # ngày mượn
                if ngayMuon:
                    df_old.iloc[i, 12] = (
                        "'" + safe_format(ngayMuon)
                    )

                # trạng thái
                df_old.iloc[i, 14] = df_new.iloc[rNew, 15]

                # ngày trả
                if ngayTra:
                    df_old.iloc[i, 18] = safe_format(ngayTra)

                # =========================================
                # LỊCH SỬ GIA HẠN
                # =========================================
                lichSuCu = str(
                    df_old.iloc[i, 20]
                ).replace("'", "")

                lichSuCu = fix_serial_date(lichSuCu)

                arr = [
                    x.strip()
                    for x in lichSuCu.split(";")
                    if x.strip()
                ]

                # thêm ngày gia hạn mới
                if (
                    ngayGiaHan
                    and ngayMuon
                    and safe_format(ngayGiaHan)
                    != safe_format(ngayMuon)
                ):

                    val = safe_format(ngayGiaHan)

                    if val not in arr:
                        arr.append(val)

                # giữ toàn bộ lịch sử
                lichSuMoi = "; ".join(arr)

                df_old.iloc[i, 20] = (
                    "'" + lichSuMoi
                    if lichSuMoi else ""
                )

                # số lần gia hạn
                df_old.iloc[i, 19] = count_gia_han(
                    lichSuMoi
                )

                # remove processed
                dict_new.pop(colID, None)

        # =================================================
        # ADD NEW ROWS
        # =================================================
        new_rows = []

        for colID, rNew in dict_new.items():

            row = [""] * 28

            row[1] = df_new.iloc[rNew, 3]
            row[2] = df_new.iloc[rNew, 4]
            row[3] = df_new.iloc[rNew, 5]
            row[4] = df_new.iloc[rNew, 6]

            row[9] = df_new.iloc[rNew, 11]

            # =============================================
            # DATE
            # =============================================
            ngayMuon = parse_date(
                df_new.iloc[rNew, 14]
            )

            ngayGiaHan = parse_date(
                df_new.iloc[rNew, 20]
            )

            ngayTra = parse_date(
                df_new.iloc[rNew, 19]
            )

            if ngayMuon:
                row[12] = "'" + safe_format(ngayMuon)

            row[14] = df_new.iloc[rNew, 15]

            if ngayTra:
                row[18] = safe_format(ngayTra)

            # =============================================
            # GIA HAN
            # =============================================
            if (
                ngayGiaHan
                and ngayMuon
                and safe_format(ngayGiaHan)
                != safe_format(ngayMuon)
            ):
                row[20] = "'" + safe_format(ngayGiaHan)

            row[19] = count_gia_han(
                str(row[20]).replace("'", "")
            )

            # =============================================
            # MÃ
            # =============================================
            valX = clean_key(
                df_new.iloc[rNew, 23]
            )

            if str(valX).isdigit():
                row[24] = "'" + str(valX).zfill(10)
            else:
                row[24] = valX

            new_rows.append(row)

        # append
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
        if path_map:

            df_map = pd.read_excel(path_map).fillna("")

            df_map.columns = range(df_map.shape[1])

            dict_map = {}

            for i in range(len(df_map)):

                key = clean_key(df_map.iloc[i, 2])

                if key:
                    dict_map[key] = i

            for i in range(len(df_old)):

                key = clean_key(df_old.iloc[i, 1])

                if key in dict_map:

                    rMap = dict_map[key]

                    df_old.iloc[i, 26] = df_map.iloc[rMap, 3]
                    df_old.iloc[i, 25] = df_map.iloc[rMap, 4]

        # =================================================
        # QUÁ HẠN
        # =================================================
        today = datetime.today()

        for i in range(len(df_old)):

            d = parse_date(df_old.iloc[i, 18])

            if d and d < today:
                df_old.iloc[i, 27] = "Qua han"
            else:
                df_old.iloc[i, 27] = "Chua qua han"

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
        tmp_output = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        )

        output_path = tmp_output.name

        tmp_output.close()

        temp_files.append(output_path)

        df_old.to_excel(
            output_path,
            index=False,
            engine="openpyxl"
        )

        # =================================================
        # HIDE COLUMNS
        # =================================================
        wb = load_workbook(output_path)

        ws = wb.active

        cols_hide = [
            "A",
            "G",
            "H",
            "I",
            "K",
            "L",
            "N",
            "P",
            "Q",
            "R",
            "V",
            "W",
            "X"
        ]

        for col in cols_hide:
            ws.column_dimensions[col].hidden = True

        wb.save(output_path)

        # =================================================
        # FILE -> BASE64
        # =================================================
        with open(output_path, "rb") as f:

            result_base64 = base64.b64encode(
                f.read()
            ).decode("utf-8")

        # =================================================
        # RETURN
        # =================================================
        return jsonify({
            "success": True,
            "file_name": "result.xlsx",
            "file_base64": result_base64
        })

    except Exception as e:

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

    finally:

        # cleanup
        for f in temp_files:

            try:
                if f and os.path.exists(f):
                    os.remove(f)
            except:
                pass


# =========================================================
# START
# =========================================================
if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5001,
        debug=True
    )
