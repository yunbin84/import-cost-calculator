from __future__ import annotations

import io
import json
import mimetypes
import os
import re
import warnings
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

warnings.filterwarnings("ignore", category=DeprecationWarning)
import cgi

import openpyxl
import pdfplumber


ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
DEFAULT_PORT = 8765
MAX_UPLOAD_BYTES = 25 * 1024 * 1024


def _num(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("¥", "").replace("₩", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def _text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _display_formula_or_value(cell: Any) -> Any:
    value = cell.value
    if isinstance(value, str) and value.startswith("="):
        return value
    return value


def parse_shipping_list(file_bytes: bytes, filename: str) -> dict[str, Any]:
    wb_formula = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    wb_values = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

    detail_name = "detail-A" if "detail-A" in wb_formula.sheetnames else wb_formula.sheetnames[0]
    detail_formula = wb_formula[detail_name]
    detail_values = wb_values[detail_name]

    items: list[dict[str, Any]] = []
    current_order = ""
    current_mark = ""
    current_name = ""
    current_url = ""

    for row_idx in range(6, detail_formula.max_row + 1):
        first = _text(detail_formula.cell(row_idx, 1).value)
        if first.upper() == "TOTAL":
            break

        order_no = _text(detail_formula.cell(row_idx, 1).value) or current_order
        mark = _text(detail_formula.cell(row_idx, 2).value) or current_mark
        url = _text(detail_formula.cell(row_idx, 3).value) or current_url
        name = _text(detail_formula.cell(row_idx, 5).value) or current_name
        option = _text(detail_formula.cell(row_idx, 6).value)
        quantity = _num(detail_values.cell(row_idx, 7).value)
        unit_cny = _num(detail_values.cell(row_idx, 8).value)
        inland_cny = _num(detail_values.cell(row_idx, 14).value)
        shipped_qty = _num(detail_values.cell(row_idx, 18).value) or quantity

        if order_no:
            current_order = order_no
        if mark:
            current_mark = mark
        if url:
            current_url = url
        if name:
            current_name = name

        has_cost_signal = any([quantity, unit_cny, inland_cny])
        if not has_cost_signal:
            continue

        items.append(
            {
                "id": f"item-{row_idx}",
                "row": row_idx,
                "orderNo": order_no,
                "mark": mark,
                "url": url,
                "name": name or option or f"Row {row_idx}",
                "option": option,
                "quantity": quantity,
                "unitCny": unit_cny,
                "inlandCny": inland_cny,
                "shippedQty": shipped_qty,
                "memo": _text(detail_formula.cell(row_idx, 11).value),
            }
        )

    shipping: list[dict[str, Any]] = []
    if "출고리스트" in wb_formula.sheetnames:
        ship_formula = wb_formula["출고리스트"]
        ship_values = wb_values["출고리스트"]
        last_title = ""
        last_en_title = ""
        for row_idx in range(10, ship_formula.max_row + 1):
            first = _text(ship_formula.cell(row_idx, 1).value)
            if "GRAND TOTAL" in first.upper():
                break

            title = _text(ship_formula.cell(row_idx, 3).value) or last_title
            en_title = _text(ship_formula.cell(row_idx, 4).value) or last_en_title
            if title:
                last_title = title
            if en_title:
                last_en_title = en_title

            no = _display_formula_or_value(ship_formula.cell(row_idx, 1))
            if no is None and not title:
                continue

            shipping.append(
                {
                    "row": row_idx,
                    "no": no,
                    "mark": _text(ship_formula.cell(row_idx, 2).value),
                    "title": title,
                    "englishTitle": en_title,
                    "material": _text(ship_formula.cell(row_idx, 5).value),
                    "spec": _text(ship_formula.cell(row_idx, 6).value),
                    "cartons": _num(ship_values.cell(row_idx, 7).value),
                    "pieces": _num(ship_values.cell(row_idx, 8).value),
                    "kgs": _num(ship_values.cell(row_idx, 9).value),
                    "cbm": _num(ship_values.cell(row_idx, 10).value),
                    "declarationUnit": _num(ship_values.cell(row_idx, 11).value),
                    "declarationTotal": _num(ship_values.cell(row_idx, 12).value),
                }
            )

    meta = {
        "filename": filename,
        "sheets": wb_formula.sheetnames,
        "detailSheet": detail_name,
        "shippingSheet": "출고리스트" if "출고리스트" in wb_formula.sheetnames else "",
    }

    return {"meta": meta, "items": items, "shipping": shipping}


def parse_invoice_pdf(file_bytes: bytes, filename: str) -> dict[str, Any]:
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    costs = {"tax": [], "logistics": [], "other": []}
    meta: dict[str, Any] = {"filename": filename}
    section = ""

    label_map = {
        "OCEAN FREIGHT CHARGE": "해상운임",
        "TERMINAL HANDLING": "터미널 핸들링",
        "DOCUMENT FEE": "서류 발급비",
        "C.F.S CHARGE": "CFS 비용",
        "HANDLING CHARGE": "핸들링 비용",
        "OTHER CHARGE": "기타 물류비",
        "STORAGE CHARGE": "창고료",
        "CERTIFICATE OF ORIGIN": "원산지 증명서",
        "CUSTOMS CLEARANCE": "통관수수료",
        "INLAND CHARGE": "국내 운송비",
    }

    for line in lines:
        if line.startswith("REF NO."):
            parts = line.split()
            if len(parts) >= 3:
                meta["refNo"] = parts[2]
        if line.startswith("H.B/L NO."):
            parts = line.split()
            if len(parts) >= 3:
                meta["hblNo"] = parts[2]
        if "ARRIVAL DATE" in line:
            maybe_date = line.rsplit(" ", 1)[-1]
            if maybe_date.count("-") == 2:
                meta["arrivalDate"] = maybe_date
        if "G.W/T / CBM" in line:
            meta["weightCbm"] = line.replace("G.W/T / CBM", "").strip()

        if "NON INVOICE" in line:
            section = "tax"
            continue
        if "INVOICE :" in line:
            section = "logistics"
            continue
        if line.startswith("SUB TOTAL") or line.startswith("TOTAL") or line.startswith("("):
            continue

        numbers = [int(raw.replace(",", "")) for raw in re.findall(r"\d{1,3}(?:,\d{3})+", line)]
        if not numbers:
            continue

        if line.startswith("관세"):
            costs["tax"].append({"label": "관세", "amount": numbers[-1], "source": line})
            continue
        if line.startswith("부가세"):
            costs["tax"].append({"label": "부가세", "amount": numbers[-1], "source": line})
            continue

        if section == "logistics":
            vat = 0
            amount = numbers[-1]
            if len(numbers) >= 2 and numbers[-1] < numbers[-2]:
                amount = numbers[-2]
                vat = numbers[-1]

            upper = line.upper()
            label = next((ko for key, ko in label_map.items() if upper.startswith(key)), "")
            if not label:
                label = line.split(" KRW ", 1)[0].strip()
            costs["logistics"].append({"label": label, "amount": amount, "source": line})
            if vat:
                costs["tax"].append({"label": f"{label} VAT", "amount": vat, "source": line})

    totals = {
        "tax": sum(cost["amount"] for cost in costs["tax"]),
        "logistics": sum(cost["amount"] for cost in costs["logistics"]),
        "other": sum(cost["amount"] for cost in costs["other"]),
    }
    totals["grand"] = totals["tax"] + totals["logistics"] + totals["other"]

    return {"meta": meta, "costs": costs, "totals": totals, "rawText": text[:4000]}


class Handler(SimpleHTTPRequestHandler):
    def translate_path(self, path: str) -> str:
        if path == "/" or path.startswith("/static/"):
            rel = "index.html" if path == "/" else path.removeprefix("/static/")
            return str(STATIC / rel)
        return str(ROOT / path.lstrip("/"))

    def do_GET(self) -> None:
        if self.path == "/health":
            self._json({"ok": True})
            return
        super().do_GET()

    def do_POST(self) -> None:
        if self.path not in {"/api/import", "/api/import-invoice"}:
            self.send_error(404)
            return

        content_length = int(self.headers.get("Content-Length", "0") or "0")
        if content_length > MAX_UPLOAD_BYTES:
            self.send_error(413, "uploaded file is too large")
            return

        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
            },
        )
        field = form["file"] if "file" in form else None
        if field is None or not getattr(field, "file", None):
            self.send_error(400, "file field is required")
            return

        try:
            file_bytes = field.file.read()
            filename = field.filename or "upload"
            if self.path == "/api/import-invoice":
                payload = parse_invoice_pdf(file_bytes, filename)
            else:
                payload = parse_shipping_list(file_bytes, filename)
        except Exception as exc:  # pragma: no cover - displayed to the local user
            self._json({"error": str(exc)}, status=500)
            return
        self._json(payload)

    def end_headers(self) -> None:
        if self.path.endswith(".js"):
            self.send_header("Content-Type", "text/javascript; charset=utf-8")
        super().end_headers()

    def guess_type(self, path: str) -> str:
        if path.endswith(".js"):
            return "text/javascript"
        return mimetypes.guess_type(path)[0] or "application/octet-stream"

    def _json(self, body: dict[str, Any], status: int = 200) -> None:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main() -> None:
    port = int(os.environ.get("PORT", DEFAULT_PORT))
    host = os.environ.get("HOST", "0.0.0.0")
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"수입원가 계산기: http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
