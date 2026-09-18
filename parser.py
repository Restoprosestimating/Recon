"""Parse an Xactimate Internal TAM PDF into structured estimate data."""
from __future__ import annotations

import re
from dataclasses import dataclass, field, asdict


REQUIRED_SECTIONS = [
    ("coversheet", re.compile(r"(Insured:|Property:|Claim Number|Price List)", re.I)),
    ("line_item_detail", re.compile(r"(DESCRIPTION|QUAN\s*UNIT|UNIT PRICE|LBR\.\s*RATE)", re.I)),
    ("summary", re.compile(r"(Replacement Cost Value|Net Claim|Line Item Total)", re.I)),
    ("recap_by_room", re.compile(r"Recap by Room", re.I)),
    ("recap_by_category", re.compile(r"Recap by Category", re.I)),
    ("labor_breakdown", re.compile(r"Labor[\s-]*Breakdown", re.I)),
    ("sketch_or_areas", re.compile(r"(Main Level|Grand Total Areas|Sketch)", re.I)),
]

PRINT_SELECTION_NAMES = [
    "Coversheet",
    "Line item detail",
    "Summary",
    "Recap by room",
    "Recap by category",
    "Labor breakdown",
    "Sketch / Grand total areas",
]


@dataclass
class LaborRow:
    code: str
    description: str
    xact_rate: float
    hours: float
    amount: float
    role: str


@dataclass
class ParseResult:
    ok: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    found_sections: list[str] = field(default_factory=list)
    missing_sections: list[str] = field(default_factory=list)
    property_line: str = ""
    insured: str = ""
    estimate_id: str = ""
    price_list: str = ""
    line_item_total: float = 0.0
    labor_direct: float = 0.0
    labor_burden: float = 0.0
    labor_loaded: float = 0.0
    materials: float = 0.0
    equipment: float = 0.0
    rcv: float = 0.0
    tax_cleaning_matl: float = 0.0
    tax_material: float = 0.0
    tax_cleaning_total: float = 0.0
    labor_rows: list[LaborRow] = field(default_factory=list)
    raw_text_preview: str = ""

    def to_dict(self):
        d = asdict(self)
        return d


def _money(s: str) -> float:
    s = s.replace(",", "").replace("$", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def extract_text(pdf_path: str) -> str:
    try:
        import pdfplumber
        parts = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                parts.append(page.extract_text() or "")
        return "\n".join(parts)
    except Exception:
        from pypdf import PdfReader
        reader = PdfReader(pdf_path)
        return "\n".join((p.extract_text() or "") for p in reader.pages)


def detect_sections(text: str) -> tuple[list[str], list[str]]:
    found, missing = [], []
    for name, pat in REQUIRED_SECTIONS:
        if pat.search(text):
            found.append(name)
        else:
            missing.append(name)
    return found, missing


def looks_like_internal_tam(text: str) -> list[str]:
    errs = []
    if not re.search(r"(HOURS|LBR\.?\s*RATE|Labor[\s-]*Breakdown)", text, re.I):
        errs.append(
            "This PDF does not look like an Internal TAM report. "
            "Hours / labor-rate columns or a Labor-Breakdown recap were not found."
        )
    if re.search(r"\bFinal Draft\b", text) and not re.search(r"Labor[\s-]*Breakdown", text, re.I):
        errs.append("This appears to be a Final Draft (or similar) estimate, not Internal TAM.")
    return errs


def parse_header(text: str, result: ParseResult) -> None:
    m = re.search(r"Property:\s*(.+)", text)
    if m:
        result.property_line = m.group(1).split("\n")[0].strip()
    m = re.search(r"Insured:\s*(.+)", text)
    if m:
        result.insured = m.group(1).split("\n")[0].strip()
    m = re.search(r"Estimate:\s*([A-Za-z0-9._-]+)", text)
    if m:
        result.estimate_id = m.group(1).strip()
    else:
        m = re.search(r"(20\d{2}-\d{2}-\d{2}-\d{4})", text)
        if m:
            result.estimate_id = m.group(1)
    m = re.search(r"Price List:\s*(\S+)", text)
    if m:
        result.price_list = m.group(1).strip()


def parse_totals(text: str, result: ParseResult) -> None:
    def grab(patterns):
        for p in patterns:
            m = re.search(p, text, re.I)
            if m:
                return _money(m.group(1))
        return 0.0

    result.line_item_total = grab([
        r"Line Item Totals?:[^\n]*?([\d,]+\.\d{2})",
        r"Line Item Total\s+([\d,]+\.\d{2})",
        r"Subtotal of Areas\s+[\d.]+%\s+([\d,]+\.\d{2})",
    ])
    result.rcv = grab([
        r"Replacement Cost Value\s+\$?([\d,]+\.\d{2})",
        r"Net Claim\s+\$?([\d,]+\.\d{2})",
    ])
    result.tax_cleaning_matl = grab([r"Cleaning Matl Tax\s+([\d,]+\.\d{2})"])
    result.tax_material = grab([r"Material Sales Tax\s+([\d,]+\.\d{2})"])
    result.tax_cleaning_total = grab([r"Cleaning Total Tax\s+([\d,]+\.\d{2})"])

    # Page-26 style rolled totals: Labor / Burden / Material / Equipment
    m = re.search(
        r"Line Item Totals:[^\n]*\n[^\n]*?([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})",
        text,
    )
    if m:
        # hours, total, labor, burden, material, equipment — layout varies
        pass

    m = re.search(
        r"Line Item Totals:.*?([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})",
        text,
        re.S,
    )
    # Fallback: recap-by-category footer style numbers near end
    # Labor / Lbr.Burden / Material / Equipment columns in category recap
    labor_m = re.search(r"\nLabor\s+Lbr\.?\s*Burden\s+Material\s+Equipment", text)
    # Direct from labor breakdown total
    m = re.search(r"Labor[\s-]*Breakdown.*?Total\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})", text, re.S | re.I)
    m2 = re.search(r"Labor[\s-]*Breakdown[\s\S]{0,2500}?Total\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})", text, re.I)
    if m2:
        # hours, amount (loaded)
        result.labor_loaded = _money(m2.group(2)) if _money(m2.group(2)) > 100 else _money(m2.group(1))

    # Materials / equipment from recap by category is unreliable without columns.
    # Pull from "Line Item Totals" trailing numbers if present.
    m = re.search(
        r"Line Item Totals:[^\n]*\n\s*([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})",
        text,
    )
    if m:
        nums = [_money(g) for g in m.groups()]
        # observed layout: hours, grand, labor, burden, material, equipment
        if nums[2] > nums[4]:
            result.labor_direct = nums[2]
            result.labor_burden = nums[3]
            result.materials = nums[4]
            result.equipment = nums[5]
            if not result.labor_loaded:
                result.labor_loaded = result.labor_direct + result.labor_burden


def parse_labor_breakdown(text: str, result: ParseResult) -> None:
    # Codes commonly used in Xactimate labor recap
    block_m = re.search(r"Labor[\s-]*Breakdown(.*?)(?:Grand Total Areas|Recap of Taxes|Main Level|Page:\s+\d+\s*$)", text, re.S | re.I)
    block = block_m.group(1) if block_m else text

    # Pattern: CODE  Description...  rate  hours  amount
    # Description can have spaces. Codes are like CLN, CLN-S, HMRT, CARP-FNC
    row_re = re.compile(
        r"\b([A-Z]{2,6}(?:-[A-Z]{1,4})?)\s+([A-Za-z][A-Za-z0-9 /,&().+-]+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})"
    )
    SUPER_CODES = {"CLN-S", "SUP", "SUPV", "PM"}
    seen = set()
    for m in row_re.finditer(block):
        code = m.group(1).strip()
        desc = re.sub(r"\s+", " ", m.group(2)).strip(" -")
        # Filter junk: code must look like a labor code, rate in plausible range
        rate = _money(m.group(3))
        a = _money(m.group(4))
        b = _money(m.group(5))
        # hours vs amount: hours usually < 5000, amount usually > hours * 20
        if a < 5000 and b >= a:
            hours, amount = a, b
        elif b < 5000 and a >= b:
            hours, amount = b, a
        else:
            hours, amount = a, b
        if not (15 <= rate <= 250):
            continue
        if hours <= 0 or hours > 20000:
            continue
        if code in seen:
            continue
        # skip if description is too short / looks like a header
        if len(desc) < 4:
            continue
        if code in {"TOTAL", "CODE", "ITEMS"}:
            continue
        seen.add(code)
        role = "Supervisor" if code in SUPER_CODES or re.search(r"supervis", desc, re.I) else "Worker"
        result.labor_rows.append(LaborRow(code, desc, rate, hours, amount, role))

    if result.labor_rows:
        result.labor_loaded = result.labor_loaded or sum(r.amount for r in result.labor_rows)
        if not result.labor_direct:
            result.labor_direct = sum(r.hours * r.xact_rate for r in result.labor_rows)


def parse_internal_tam(pdf_path: str) -> ParseResult:
    result = ParseResult(ok=False)
    try:
        text = extract_text(pdf_path)
    except Exception as e:
        result.errors.append(f"Could not read PDF: {e}")
        return result

    result.raw_text_preview = text[:1500]
    if len(text.strip()) < 200:
        result.errors.append("PDF has almost no extractable text. Scan-only PDFs are not supported.")
        return result

    result.errors.extend(looks_like_internal_tam(text))
    found, missing = detect_sections(text)
    result.found_sections = found
    result.missing_sections = missing

    # Labor breakdown is mandatory for recon
    if "labor_breakdown" in missing:
        result.errors.append(
            "Labor breakdown is missing. In Xactimate go to Documents > Reports, "
            "choose Report type = Internal TAM, and check ALL print selections "
            "(including Labor breakdown)."
        )

    parse_header(text, result)
    parse_totals(text, result)
    parse_labor_breakdown(text, result)

    if not result.labor_rows:
        result.errors.append(
            "No labor-breakdown rows could be parsed. Confirm Report type is Internal TAM "
            "and the Labor breakdown print selection is checked."
        )
    if not result.rcv and not result.line_item_total:
        result.warnings.append("Could not find Replacement Cost Value / Line Item Total. Enter negotiated price manually.")

    # Materials/equipment fallback from category-like lines
    if result.materials == 0:
        mats = re.findall(r"Material\s+([\d,]+\.\d{2})", text)
        # last large one is often a total — skip
    result.ok = len(result.errors) == 0
    return result
