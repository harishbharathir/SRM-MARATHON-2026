import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.sax.saxutils import escape

import openpyxl
from flask import Flask, Response, jsonify, render_template, request

BASE_DIR = Path(__file__).resolve().parent
SHEETS = [s.strip() for s in os.environ.get("EXCEL_SHEETS", "Male S+E,Female S+E").split(",") if s.strip()]

# ISO/IEC 15417 Code 128 patterns (values 0..106)
CODE128_PATTERNS = [
    "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312", "132212", "221213",
    "221312", "231212", "112232", "122132", "122231", "113222", "123122", "123221", "223211", "221132",
    "221231", "213212", "223112", "312131", "311222", "321122", "321221", "312212", "322112", "322211",
    "212123", "212321", "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313",
    "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121", "313121", "211331",
    "231131", "213113", "213311", "213131", "311123", "311321", "331121", "312113", "312311", "332111",
    "314111", "221411", "431111", "111224", "111422", "121124", "121421", "141122", "141221", "112214",
    "112412", "122114", "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111",
    "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112", "421211", "212141",
    "214121", "412121", "111143", "111341", "131141", "114113", "114311", "411113", "411311", "113141",
    "114131", "311141", "411131", "211412", "211214", "211232", "2331112",
]


def clean_str(val) -> str:
    """Convert cell value to a trimmed string, preserving integer IDs without .0."""
    if val is None:
        return ""
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    return str(val).strip()


def clean_phone_digits(val) -> str:
    """Keep digits only from a phone input or cell value."""
    return re.sub(r"\D", "", clean_str(val))


def normalize_10(digits: str) -> str:
    """Normalize leading 0, 91, or 091 country prefixes to a 10-digit Indian mobile number."""
    if len(digits) == 10:
        return digits
    if len(digits) == 11 and digits.startswith("0"):
        return digits[1:]
    if len(digits) == 12 and digits.startswith("91"):
        return digits[2:]
    if len(digits) == 13 and digits.startswith("091"):
        return digits[3:]
    return digits


def clean_tshirt(val) -> str:
    """Format T-shirt size spacing consistently."""
    s = clean_str(val)
    s = re.sub(r"\s*-\s*", " - ", s)
    return re.sub(r"\s+", " ", s)


def resolve_excel_path() -> Optional[Path]:
    """Locate the participant Excel workbook across environment and Render Secret File paths."""
    candidates: List[Path] = []
    env_path = os.environ.get("EXCEL_PATH", "").strip()
    if env_path:
        candidates.append(Path(env_path))
    candidates.extend([
        BASE_DIR / "SRM RUN 6.0 FINAL 1.xlsx",
        Path("/etc/secrets/SRM RUN 6.0 FINAL 1.xlsx"),
        Path("/etc/secrets/SRM_RUN_6_0_FINAL_1.xlsx"),
    ])
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def _add_phone_key(index: Dict[str, List[dict]], key: str, participant: dict) -> None:
    if not key:
        return
    bucket = index.setdefault(key, [])
    if not any(item["bib"] == participant["bib"] for item in bucket):
        bucket.append(participant)


def load_participants() -> Tuple[List[dict], Dict[str, List[dict]], Dict[str, dict]]:
    """Read participants from Excel once and build O(1) lookup indexes by phone and Bib ID."""
    excel_path = resolve_excel_path()
    if excel_path is None:
        return [], {}, {}

    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    participants: List[dict] = []
    phone_index: Dict[str, List[dict]] = {}
    bib_index: Dict[str, dict] = {}

    try:
        for sheet in SHEETS:
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            headers = None
            for idx, row in enumerate(ws.iter_rows(values_only=True)):
                if idx == 0:
                    headers = list(row)
                    continue
                if not any(cell is not None for cell in row):
                    if idx > 1000:
                        break
                    continue

                d = dict(zip(headers, row))
                bib = clean_str(d.get("Bib ID"))
                if not bib:
                    continue

                raw_digits = clean_phone_digits(d.get("Phone Number"))
                norm_phone = normalize_10(raw_digits)

                er_digits = clean_phone_digits(d.get("ER Mobile Number"))
                er_norm = normalize_10(er_digits) if er_digits else ""

                age_str = ""
                age_val = d.get("Age")
                if age_val not in (None, ""):
                    try:
                        age_int = int(round(float(age_val)))
                        if age_int > 0:
                            age_str = str(age_int)
                    except (TypeError, ValueError):
                        pass

                participant = {
                    "bib": bib,
                    "name": clean_str(d.get("Full Name")),
                    "phone": norm_phone,
                    "age": age_str,
                    "dob": clean_str(d.get("Date of Birth")),
                    "gender": clean_str(d.get("Gender")),
                    "pos": clean_str(d.get("Position")),
                    "series": clean_str(d.get("Bib Series")),
                    "group": clean_str(d.get("Series Group")),
                    "inst": clean_str(d.get("Institutions")),
                    "dept": clean_str(d.get("Department")),
                    "year": clean_str(d.get("Designation / Studying Year")),
                    "blood": clean_str(d.get("Blood Group")),
                    "tshirt": clean_tshirt(d.get("T Shirt Size")),
                    "er": er_norm,
                    "transport": clean_str(d.get("Transport Needed")) or "Not required",
                    "email": clean_str(d.get("Email Address")),
                    "bookingId": clean_str(d.get("Booking ID")),
                    "status": clean_str(d.get("Booking status")),
                    "payStatus": clean_str(d.get("Payment status")),
                    "sno": clean_str(d.get("S.No")),
                }

                participants.append(participant)
                bib_index[bib.upper()] = participant

                aliases: List[str] = []
                if raw_digits and raw_digits != norm_phone:
                    aliases.append(raw_digits)
                if bib == "SF321" and raw_digits == "73589880008":
                    aliases.append("7358988008")
                elif bib == "SF673" and raw_digits == "81226388970":
                    aliases.append("8122638970")
                elif bib == "SF417" and raw_digits == "632721815824" and er_norm:
                    aliases.append(er_norm)

                _add_phone_key(phone_index, norm_phone, participant)
                for alias in aliases:
                    _add_phone_key(phone_index, alias, participant)
    finally:
        wb.close()

    return participants, phone_index, bib_index


def encode_code128_b(text: str) -> Tuple[List[Tuple[bool, int]], int]:
    """Encode an ASCII string using ISO/IEC 15417 Code 128 Set B."""
    codes = [104]  # Start B
    checksum = 104
    for pos, ch in enumerate(text, start=1):
        ascii_code = ord(ch)
        if ascii_code < 32 or ascii_code > 127:
            raise ValueError(f"Unsupported Code 128 character: {ch!r}")
        val = ascii_code - 32
        codes.append(val)
        checksum += pos * val
    codes.append(checksum % 103)
    codes.append(106)  # Stop

    runs: List[Tuple[bool, int]] = []
    total_modules = 0
    for code in codes:
        pat = CODE128_PATTERNS[code]
        for j, ch in enumerate(pat):
            units = int(ch)
            runs.append((j % 2 == 0, units))
            total_modules += units
    return runs, total_modules


def render_code128_svg(bib: str) -> str:
    """Render a crisp, high-contrast Code 128 SVG with quiet zones and Bib caption."""
    runs, total_modules = encode_code128_b(bib)
    mod_w = 3
    quiet = 12 * mod_w
    bars_w = total_modules * mod_w
    svg_w = bars_w + quiet * 2
    bar_y = 14
    bar_h = 84
    svg_h = 132

    rects = []
    cursor_x = quiet
    for is_bar, units in runs:
        rw = units * mod_w
        if is_bar:
            rects.append(f'<rect x="{cursor_x}" y="{bar_y}" width="{rw}" height="{bar_h}" fill="#000000"/>')
        cursor_x += rw

    safe_bib = escape(bib)
    bars_markup = "".join(rects)
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{svg_w}" height="{svg_h}" '
        f'viewBox="0 0 {svg_w} {svg_h}" role="img" aria-label="Barcode {safe_bib}">'
        f'<rect width="{svg_w}" height="{svg_h}" fill="#ffffff"/>'
        f'{bars_markup}'
        f'<text x="{svg_w / 2:.1f}" y="{bar_y + bar_h + 24}" text-anchor="middle" '
        f'font-family="Consolas, Courier New, monospace" font-size="18" font-weight="700" fill="#000000">'
        f'{safe_bib}</text></svg>'
    )


app = Flask(__name__)
PARTICIPANTS, PHONE_INDEX, BIB_INDEX = load_participants()


@app.after_request
def apply_security_headers(resp: Response) -> Response:
    """Attach basic content security and no-store caching for participant API lookups."""
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    if request.path.startswith("/api/search"):
        resp.headers["Cache-Control"] = "no-store"
    cors_origin = os.environ.get("CORS_ALLOW_ORIGIN", "").strip()
    if cors_origin:
        resp.headers["Access-Control-Allow-Origin"] = cors_origin
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


@app.route("/", methods=["GET"])
@app.route("/index.html", methods=["GET"])
def home():
    """Serve the TSRM Run 2026 participant portal."""
    return render_template("index.html")


@app.route("/api/health", methods=["GET"])
def health():
    """Health check endpoint used by Render deployment monitoring."""
    return jsonify({
        "ok": len(PARTICIPANTS) > 0,
        "status": "healthy" if PARTICIPANTS else "dataset_missing",
        "event": "TSRM Run 2026",
        "eventDate": "2026-11-27",
        "participantsLoaded": len(PARTICIPANTS),
    }), (200 if PARTICIPANTS else 503)


@app.route("/api/search", methods=["POST", "GET"])
def search_participants():
    """Look up registered participant(s) matching a mobile number."""
    if request.method == "POST":
        payload = request.get_json(silent=True) or {}
        raw_phone = clean_str(payload.get("phone"))
    else:
        raw_phone = clean_str(request.args.get("phone"))

    digits = re.sub(r"\D", "", raw_phone)
    if not digits:
        return jsonify({"ok": False, "error": "Please enter your registered mobile number."}), 400

    ten = normalize_10(digits)
    matches = PHONE_INDEX.get(digits) or PHONE_INDEX.get(ten) or []

    if len(digits) < 10 or (len(ten) != 10 and not matches):
        return jsonify({"ok": False, "error": "Please enter a valid 10-digit mobile number."}), 400

    if not matches:
        return jsonify({
            "ok": False,
            "error": (
                f"No registration found for mobile number {ten}. "
                "Please check the number used during registration or contact the TSRM Run helpdesk."
            ),
        }), 404

    return jsonify({
        "ok": True,
        "query": ten,
        "count": len(matches),
        "participants": matches,
    })


@app.route("/api/barcode/<bib>.svg", methods=["GET"])
@app.route("/api/barcode/<bib>", methods=["GET"])
def barcode_svg(bib: str):
    """Serve a crisp server-generated Code 128 SVG barcode for a registered Bib ID."""
    clean_bib = clean_str(bib).upper()
    participant = BIB_INDEX.get(clean_bib)
    if not participant:
        return jsonify({"ok": False, "error": f"Unknown Bib ID: {clean_bib}"}), 404
    svg = render_code128_svg(participant["bib"])
    return Response(svg, mimetype="image/svg+xml")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    debug = os.environ.get("FLASK_DEBUG", "").lower() in {"1", "true", "yes"}
    app.run(host="0.0.0.0", port=port, debug=debug)
