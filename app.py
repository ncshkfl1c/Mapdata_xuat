from flask import Flask, request, jsonify
from openpyxl import load_workbook
from datetime import datetime
from io import BytesIO
import tempfile
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

    if isinstance(val, datetime):
        return val

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
    # EXCEL SERIAL
    # =====================================================
    try:

        num = float(s)

        if num > 30000:

            return datetime(
                1899,
                12,
                30
            ) + timedelta(days=num)

    except:
        pass

    return None


def safe_format(d):

    if not d:
        return ""

    return d.strftime("%d/%m/%Y")


# =========================================================
# COUNT
# =========================================================
def count_gia_han(text):

    if not text:
        return ""

    arr = str(text).split(";")

    count = 0

    for x in arr:

        if str(x).strip():
            count += 1

    return str(count)


# =========================================================
# HISTORY
# =========================================================
def normalize_history(text):

    if not text:
        return ""

    text = str(text).replace("'", "").strip()

    arr = []

    for x in text.split(";"):

        x = x.strip()

        if not x:
            continue

        d = parse_date(x)

        if d:
            val = safe_format(d)
        else:
            val = x

        if val not in arr:
            arr.append(val)

    return "; ".join(arr)


# =========================================================
# HOME
# =========================================================
@app.route("/", methods=["GET"])
def home():

    return jsonify({
        "success": True,
        "message": "API RUNNING"
    })


# =========================================================
# API
# =========================================================
@app.route("/dongbo", methods=["POST"])
def dongbo():

    try:

        # =================================================
        # JSON
        # =================================================
        data = request.get_json()

        if not data:

            return jsonify({
                "success": False,
                "error": "Missing JSON"
            }), 400

        # =================================================
        # FILES
        # =================================================
        file_old = b64_to_file(
            data.get("file_old")
        )

        file_new = b64_to_file(
            data.get("file_new")
        )

        file_map = b64_to_file(
            data.get("file_map")
        )

        if not file_old:

            return jsonify({
                "success": False,
                "error": "Missing file_old"
            }), 400

        if not file_new:

            return jsonify({
                "success": False,
                "error": "Missing file_new"
            }), 400

        # =================================================
        # LOAD WORKBOOK
        # =================================================
        wb_old = load_workbook(file_old)
        ws_old = wb_old.active

        wb_new = load_workbook(file_new)
        ws_new = wb_new.active

        wb_map = None
        ws_map = None

        if file_map:

            wb_map = load_workbook(file_map)
            ws_map = wb_map.active

        # =================================================
        # CONSTANT
        # =================================================
        START_NEW = 10
        START_OLD = 2

        # =================================================
        # DICT NEW
        # =================================================
        dict_all = {}
        dict_new_only = {}

        last_new = ws_new.max_row

        for r in range(START_NEW, last_new + 1):

            key = clean_key(
                ws_new[f"L{r}"].value
            )

            if key:

                dict_all[key] = r
                dict_new_only[key] = r

        # =================================================
        # UPDATE OLD
        # =================================================
        last_old = ws_old.max_row

        rows_delete = []

        for r in range(
            last_old,
            START_OLD - 1,
            -1
        ):

            colID = clean_key(
                ws_old[f"J{r}"].value
            )

            if colID in dict_all:

                rNew = dict_all[colID]

                # =========================================
                # BASIC
                # =========================================
                ws_old[f"B{r}"] = ws_new[f"D{rNew}"].value
                ws_old[f"C{r}"] = ws_new[f"E{rNew}"].value
                ws_old[f"D{r}"] = ws_new[f"F{rNew}"].value
                ws_old[f"E{r}"] = ws_new[f"G{rNew}"].value

                # =========================================
                # M
                # =========================================
                ngayMuon = parse_date(
                    ws_new[f"O{rNew}"].value
                )

                if ngayMuon:

                    ws_old[f"M{r}"] = safe_format(
                        ngayMuon
                    )

                else:

                    ws_old[f"M{r}"] = (
                        ws_new[f"O{rNew}"].value
                    )

                # =========================================
                # O
                # =========================================
                ws_old[f"O{r}"] = (
                    ws_new[f"P{rNew}"].value
                )

                # =========================================
                # S
                # =========================================
                ngayTra = parse_date(
                    ws_new[f"T{rNew}"].value
                )

                if ngayTra:

                    ws_old[f"S{r}"] = safe_format(
                        ngayTra
                    )

                else:

                    ws_old[f"S{r}"] = (
                        ws_new[f"T{rNew}"].value
                    )

                # =========================================
                # U
                # =========================================
                ngayGiaHan = parse_date(
                    ws_new[f"U{rNew}"].value
                )

                lichSuCu = normalize_history(
                    ws_old[f"U{r}"].value
                )

                arr = []

                if lichSuCu:

                    arr = [
                        x.strip()
                        for x in lichSuCu.split(";")
                        if x.strip()
                    ]

                if (
                    ngayGiaHan and
                    ngayMuon and
                    safe_format(ngayGiaHan)
                    !=
                    safe_format(ngayMuon)
                ):

                    val = safe_format(
                        ngayGiaHan
                    )

                    if val not in arr:
                        arr.append(val)

                lichSuCu = "; ".join(arr)

                ws_old[f"U{r}"] = lichSuCu

                # =========================================
                # T
                # =========================================
                ws_old[f"T{r}"] = count_gia_han(
                    lichSuCu
                )

                # =========================================
                # REMOVE
                # =========================================
                dict_new_only.pop(
                    colID,
                    None
                )

            else:

                rows_delete.append(r)

        # =================================================
        # DELETE ROW
        # =================================================
        for r in rows_delete:

            ws_old.delete_rows(r)

        # =================================================
        # ADD NEW
        # =================================================
        for colID, rNew in dict_new_only.items():

            newRow = ws_old.max_row + 1

            # =============================================
            # BASIC
            # =============================================
            ws_old[f"B{newRow}"] = ws_new[f"D{rNew}"].value
            ws_old[f"C{newRow}"] = ws_new[f"E{rNew}"].value
            ws_old[f"D{newRow}"] = ws_new[f"F{rNew}"].value
            ws_old[f"E{newRow}"] = ws_new[f"G{rNew}"].value
            ws_old[f"J{newRow}"] = ws_new[f"L{rNew}"].value

            # =============================================
            # M
            # =============================================
            ngayMuon = parse_date(
                ws_new[f"O{rNew}"].value
            )

            if ngayMuon:

                ws_old[f"M{newRow}"] = safe_format(
                    ngayMuon
                )

            # =============================================
            # O
            # =============================================
            ws_old[f"O{newRow}"] = (
                ws_new[f"P{rNew}"].value
            )

            # =============================================
            # S
            # =============================================
            ngayTra = parse_date(
                ws_new[f"T{rNew}"].value
            )

            if ngayTra:

                ws_old[f"S{newRow}"] = safe_format(
                    ngayTra
                )

            # =============================================
            # U
            # =============================================
            ngayGiaHan = parse_date(
                ws_new[f"U{rNew}"].value
            )

            if (
                ngayGiaHan and
                ngayMuon and
                safe_format(ngayGiaHan)
                !=
                safe_format(ngayMuon)
            ):

                ws_old[f"U{newRow}"] = safe_format(
                    ngayGiaHan
                )

            # =============================================
            # X -> Y
            # =============================================
            valX = str(
                ws_new[f"X{rNew}"].value
            )

            if valX.isdigit():

                ws_old[f"Y{newRow}"] = (
                    valX.zfill(10)
                )

            else:

                ws_old[f"Y{newRow}"] = valX

            # =============================================
            # T
            # =============================================
            ws_old[f"T{newRow}"] = count_gia_han(
                ws_old[f"U{newRow}"].value
            )

        # =================================================
        # MAP
        # =================================================
        if ws_map:

            dict_map = {}

            last_map = ws_map.max_row

            for r in range(2, last_map + 1):

                key = clean_key(
                    ws_map[f"C{r}"].value
                )

                if key:

                    dict_map[key] = r

            last_old = ws_old.max_row

            for r in range(2, last_old + 1):

                key = clean_key(
                    ws_old[f"B{r}"].value
                )

                if key in dict_map:

                    rMap = dict_map[key]

                    ws_old[f"AA{r}"] = (
                        ws_map[f"D{rMap}"].value
                    )

                    ws_old[f"Z{r}"] = (
                        ws_map[f"E{rMap}"].value
                    )

        # =================================================
        # STATUS
        # =================================================
        last_old = ws_old.max_row

        ws_old["AB1"] = "Tình trạng"

        today = datetime.today()

        for r in range(2, last_old + 1):

            d = parse_date(
                ws_old[f"S{r}"].value
            )

            if d and d < today:

                ws_old[f"AB{r}"] = "Qua han"

            else:

                ws_old[f"AB{r}"] = "Chua qua han"

        # =================================================
        # HIDE COL
        # =================================================
        cols_hide = [
            "A", "G", "H", "I",
            "K", "L", "N", "P",
            "Q", "R", "V", "W", "X"
        ]

        for col in cols_hide:

            ws_old.column_dimensions[
                col
            ].hidden = True

        # =================================================
        # SAVE
        # =================================================
        tmp = tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".xlsx"
        )

        path = tmp.name

        tmp.close()

        wb_old.save(path)

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
