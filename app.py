from flask import Flask, request, jsonify
import pandas as pd
from datetime import datetime
import tempfile
import base64
import io
from openpyxl import load_workbook

app = Flask(__name__)

# ================= CLEAN KEY =================
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


# ================= DATE =================
def parse_date(val):
    if val is None:
        return None

    s = str(val).strip()
    if s == "":
        return None

    try:
        return datetime.strptime(s, "%d/%m/%Y")
    except:
        pass

    try:
        d = pd.to_datetime(val, origin='1899-12-30', unit='D', errors='coerce')
        if pd.isna(d):
            return None
        return d.to_pydatetime()
    except:
        pass

    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return None
        return d.to_pydatetime()
    except:
        return None


def safe_format(d):
    return d.strftime("%d/%m/%Y") if d else ""


# ================= BASE64 TO DATAFRAME =================
def read_excel_base64(b64_string):
    file_bytes = base64.b64decode(b64_string)
    return pd.read_excel(io.BytesIO(file_bytes), dtype=str).fillna("")


# ================= COUNT / HISTORY =================
def fix_serial_date(text):
    if not text:
        return ""

    parts = str(text).split(";")
    result = []

    for p in parts:
        d = parse_date(p.strip())
        val = d.strftime("%d/%m/%Y") if d else p.strip()

        if val and val not in result:
            result.append(val)

    return "; ".join(result)


def count_gia_han(text):
    if not text:
        return ""
    return str(len([x for x in text.split(";") if x.strip()]))


# ================= API =================
@app.route("/dongbo", methods=["POST"])
def dongbo():
    try:
        data = request.get_json()

        file_old_b64 = data["file_old"]
        file_new_b64 = data["file_new"]
        file_map_b64 = data.get("file_map")

        # ===== READ EXCEL FROM BASE64 =====
        df_old = read_excel_base64(file_old_b64)
        df_new = read_excel_base64(file_new_b64)

        if file_map_b64:
            df_map = read_excel_base64(file_map_b64)
        else:
            df_map = None

        # ===== SKIP HEADER NEW =====
        df_new = df_new.iloc[9:].reset_index(drop=True)

        # ===== FIX SIZE =====
        if df_old.shape[1] < 28:
            for i in range(df_old.shape[1], 28):
                df_old[i] = ""

        COL_OLD_ID = 9
        COL_NEW_ID = 11

        dict_all = {}
        dict_new_only = {}

        # ===== NEW MAP =====
        for i in range(len(df_new)):
            key = clean_key(df_new.iloc[i, COL_NEW_ID])
            if key:
                dict_all[key] = i
                dict_new_only[key] = i

        rows_keep = [0]

        # ===== UPDATE =====
        for i in reversed(range(1, len(df_old))):
            colID = clean_key(df_old.iloc[i, COL_OLD_ID])

            if colID in dict_all:
                rNew = dict_all[colID]

                df_old.iloc[i, 1] = df_new.iloc[rNew, 3]
                df_old.iloc[i, 2] = df_new.iloc[rNew, 4]
                df_old.iloc[i, 3] = df_new.iloc[rNew, 5]
                df_old.iloc[i, 4] = df_new.iloc[rNew, 6]

                ngayMuon = parse_date(df_new.iloc[rNew, 14])
                ngayGiaHan = parse_date(df_new.iloc[rNew, 20])

                df_old.iloc[i, 12] = "'" + safe_format(ngayMuon) if ngayMuon else df_new.iloc[rNew, 14]
                df_old.iloc[i, 14] = df_new.iloc[rNew, 15]
                df_old.iloc[i, 18] = safe_format(parse_date(df_new.iloc[rNew, 19]))

                lichSuCu = str(df_old.iloc[i, 20]).replace("'", "")
                lichSuCu = fix_serial_date(lichSuCu)

                arr = [x.strip() for x in lichSuCu.split(";") if x.strip()]

                if ngayGiaHan and ngayMuon and ngayGiaHan != ngayMuon:
                    val = safe_format(ngayGiaHan)
                    if val not in arr:
                        arr.append(val)

                lichSuCu = "; ".join(arr)

                df_old.iloc[i, 20] = "'" + lichSuCu if lichSuCu else ""
                df_old.iloc[i, 19] = count_gia_han(lichSuCu)

                dict_new_only.pop(colID, None)
                rows_keep.append(i)

        df_old = df_old.iloc[rows_keep].reset_index(drop=True)

        # ===== ADD NEW =====
        new_rows = []

        for colID, rNew in dict_new_only.items():
            row = [""] * 28

            row[1] = df_new.iloc[rNew, 3]
            row[2] = df_new.iloc[rNew, 4]
            row[3] = df_new.iloc[rNew, 5]
            row[4] = df_new.iloc[rNew, 6]
            row[9] = df_new.iloc[rNew, 11]

            ngayMuon = parse_date(df_new.iloc[rNew, 14])
            ngayGiaHan = parse_date(df_new.iloc[rNew, 20])

            if ngayMuon:
                row[12] = "'" + safe_format(ngayMuon)

            row[14] = df_new.iloc[rNew, 15]

            ngayTra = parse_date(df_new.iloc[rNew, 19])
            if ngayTra:
                row[18] = safe_format(ngayTra)

            if ngayGiaHan and ngayMuon and ngayGiaHan != ngayMuon:
                row[20] = "'" + safe_format(ngayGiaHan)

            valX = str(df_new.iloc[rNew, 23])
            row[24] = "'" + valX.zfill(10) if valX.isdigit() else valX

            row[19] = count_gia_han(str(row[20]).replace("'", ""))

            new_rows.append(row)

        if new_rows:
            df_add = pd.DataFrame(new_rows, columns=df_old.columns)
            df_old = pd.concat([df_old, df_add], ignore_index=True)

        # ===== MAP =====
        if df_map is not None:
            dict_map = {}

            for i in range(len(df_map)):
                key = clean_key(df_map.iloc[i, 2])
                if key:
                    dict_map[key] = i

            for i in range(1, len(df_old)):
                key = clean_key(df_old.iloc[i, 1])
                if key in dict_map:
                    rMap = dict_map[key]
                    df_old.iloc[i, 26] = df_map.iloc[rMap, 3]
                    df_old.iloc[i, 25] = df_map.iloc[rMap, 4]

        # ===== OVERDUE =====
        today = datetime.today()

        for i in range(len(df_old)):
            d = parse_date(df_old.iloc[i, 18])
            df_old.iloc[i, 27] = "Qua han" if d and d < today else "Chua qua han"

        # ===== EXPORT =====
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        path = tmp.name
        tmp.close()

        df_old.to_excel(path, index=False, engine="openpyxl")

        wb = load_workbook(path)
        ws = wb.active

        for col in ["A","G","H","I","K","L","N","P","Q","R","V","W","X"]:
            ws.column_dimensions[col].hidden = True

        wb.save(path)

        return jsonify({
            "file": base64.b64encode(open(path, "rb").read()).decode("utf-8")
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5001)
