from flask import Flask, request, send_file, jsonify
import pandas as pd
from datetime import datetime
import tempfile
import os
import base64
import io
from openpyxl import load_workbook

app = Flask(__name__)

# ====================== CÁC HÀM HỖ TRỢ ======================

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


def parse_date(val):
    if val is None or str(val).strip() == "":
        return None
    s = str(val).strip()

    # Thử định dạng Việt Nam
    for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"]:
        try:
            return datetime.strptime(s, fmt)
        except:
            continue

    # Excel serial date
    try:
        d = pd.to_datetime(val, origin='1899-12-30', unit='D', errors='coerce')
        if not pd.isna(d):
            return d.to_pydatetime()
    except:
        pass

    # Fallback
    try:
        d = pd.to_datetime(val, errors='coerce')
        if not pd.isna(d):
            return d.to_pydatetime()
    except:
        pass
    return None


def safe_format(d):
    return d.strftime("%d/%m/%Y") if d else ""


def fix_serial_date(text):
    if not text:
        return ""
    parts = str(text).split(";")
    result = []
    for p in parts:
        d = parse_date(p.strip())
        val = safe_format(d) if d else p.strip()
        if val and val not in result:
            result.append(val)
    return "; ".join(result)


def count_gia_han(text):
    if not text:
        return "0"
    return str(len([x for x in text.split(";") if x.strip()]))


def ensure_columns(df, total_cols):
    if df.shape[1] < total_cols:
        for i in range(df.shape[1], total_cols):
            df[i] = ""
    return df


# ====================== API NHẬN BASE64 ======================
@app.route("/dongbo", methods=["POST"])
def dongbo():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "Không nhận được dữ liệu JSON"}), 400

        file_old_b64 = data.get("file_old")
        file_new_b64 = data.get("file_new")
        file_map_b64 = data.get("file_map")  # optional

        if not file_old_b64 or not file_new_b64:
            return jsonify({"error": "Thiếu file_old hoặc file_new (Base64)"}), 400

        # Hàm chuyển Base64 sang DataFrame
        def b64_to_df(b64_string, skiprows=0):
            file_bytes = base64.b64decode(b64_string)
            df = pd.read_excel(io.BytesIO(file_bytes), dtype=str, skiprows=skiprows).fillna("")
            return df

        df_old = b64_to_df(file_old_b64)
        df_new = b64_to_df(file_new_b64, skiprows=9)   # Skip 9 dòng đầu file new

        df_old = ensure_columns(df_old, 28)

        COL_OLD_ID = 9
        COL_NEW_ID = 11

        dict_all = {}
        dict_new_only = {}

        for i in range(len(df_new)):
            key = clean_key(df_new.iloc[i, COL_NEW_ID])
            if key:
                dict_all[key] = i
                dict_new_only[key] = i

        rows_keep = [0]

        # === UPDATE EXISTING ===
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

        # === ADD NEW RECORDS ===
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

        # === MAP FILE (nếu có) ===
        if file_map_b64:
            df_map = b64_to_df(file_map_b64)
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

        # === TÌNH TRẠNG QUÁ HẠN ===
        today = datetime.today()
        for i in range(len(df_old)):
            d = parse_date(df_old.iloc[i, 18])
            df_old.iloc[i, 27] = "Qua han" if d and d < today else "Chua qua han"

        # Đặt tên cột
        cols = list(df_old.columns)
        if len(cols) > 27:
            cols[27] = "Tình trạng"
        df_old.columns = cols

        # ====================== EXPORT ======================
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        path = tmp.name
        tmp.close()

        df_old.to_excel(path, index=False, engine="openpyxl")

        # Hide columns
        wb = load_workbook(path)
        ws = wb.active
        cols_hide = ["A", "G", "H", "I", "K", "L", "N", "P", "Q", "R", "V", "W", "X"]
        for col in cols_hide:
            ws.column_dimensions[col].hidden = True
        wb.save(path)

        return send_file(path, as_attachment=True, download_name="KetQua_DongBo.xlsx")

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5001)
