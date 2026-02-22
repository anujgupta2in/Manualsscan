# extraction_utils.py
# ------------------------------------------------------------
# Robust title + type extraction utilities for ship manuals/drawings.
# Fixes:
#   - "OF Provision Refer Plant" / "OF System" by preserving V/D prefix
#   - OCR garbage (OLT/COL/DOL etc.)
#   - Prefer filename for "M(A)-xx ..." drawings
# ------------------------------------------------------------

import re
import unicodedata
from pathlib import Path

SYSTEM_KEYWORDS = [
    "ARRANGEMENT", "ARR'T", "ARR T", "ARRANGMENT",
    "E/R", "E-R", "E R",
    "WORK SHOP", "WORKSHOP", "STORE",
    "SPARE PARTS", "TOOLS LIST", "SPARE PARTS AND TOOLS LIST", "SPARE PARTS & TOOLS LIST",
    "MAIN ENGINE", "AUX", "AUX. MACHINERY", "AUX MACHINERY",
    "VALVE LIST", "V/V LIST",
    "ICCP", "I.C.C.P.",
    "PAINTING SPECIFICATION",
    "DOCKING ANALYSIS", "DOCKING PLAN",
    "ACCOMMODATION", "DECK MACHINERY",
    "SIDE THRUSTER", "STERN THRUSTER", "BOW THRUSTER",
    "FIRE FIGHTING", "FIRE APPLIANCE",
    "AIR CONDITIONING PLANT",
    "PROVISION REFRIGERATING PLANT",
    "GALLEY AND LAUNDRY EQUIPMENT",
    "SEWAGE TREATMENT PLANT",
    "MACHINERY ARRANGEMENT",
    "EQUIPMENT LIST",
    "FLOWMETER", "FLOW METER",
    "PLATE TYPE COOLER",
    "CENTRIFUGAL PUMP",
    "M.G.P.S", "M. G. P. S", "MARINE GROWTH PREVENTION SYSTEM",
    "STEERING GEAR ROOM", "STEERING GEAR", "STEERI NG GEAR",
    "RUDDER", "SHAFT", "FOUNDATION", "FRAME", "STOCK", "CONSTRUCTION", "STERN",
    "HULL", "SHELL", "EXPANSION", "BEARING",
    "NAME", "CAUTION", "PLATE", "NAME & CAUTION PLATE", "NAME AND CAUTION PLATE",
    "INSTALLATION", "DOOR PLAN", "INSULATION PLAN", "CRANE", "LIFTING BEAM",
    "V/D", "V-D", "VENDOR DRAWING",
]

LABEL_PATTERNS = {
    "Document Title": r"(?:Title|Manual for|Drawing Title|Plan Title)\s*[:\-]\s*([A-Z\s/&.\-]+)",
    "Engine Type": r"(?:Engine|M/E|A/E)\s*type\s*[:\-]?\s*([A-Z0-9/.\-]+)",
    "Ship No": r"(?:Ship|Hull|Project|H\.?NO\.?)\s*(?:no\.?|Number)\s*[:\-]?\s*([A-Z0-9/.\-]+)",
    "Vessel Name": r"(?:Vessel|Ship)\s*Name\s*[:\-]?\s*([A-Z\s]+)"
}

METADATA_PATTERNS = {
    "Document Number": r"(?:Doc|DWG|Plan)\.?\s*(?:No|Number)\s*[:\-]?\s*([A-Z0-9/.\-]+)",
    "Maker": r"(?:Maker|Manufacturer|Company)\s*[:\-]?\s*([A-Z\s,]+)",
    "Model": r"(?:Model|Type)\s*[:\-]?\s*([A-Z0-9/.\-]+)"
}

# Strong anchors only (don’t include SYSTEM alone, it causes "OF System")
ANCHOR_WORDS = [
    "V/D", "VENDOR", "DRAWING",
    "SPARE", "TOOLS", "LIST",
    "ARRANGEMENT", "VALVE", "EQUIPMENT",
    "NAME", "CAUTION", "PLATE",
    "ICCP", "PAINTING", "DOCKING",
    "FLOWMETER", "COOLER", "PUMP",
    "STEERING", "RUDDER", "FOUNDATION", "CONSTRUCTION",
    "INSTALLATION", "DOOR", "INSULATION", "CRANE", "LIFTING",
    "AIR-CON", "AIRCON", "PROVISION", "REFRIGERATING",
]

STAMP_REGEX = re.compile(
    r"(?i)\b("
    r"DATE|REV\.?|DESCRIPTION|OWN|CHKD\.?|APPD\.?|DWN\.?|"
    r"CHKD\s*BY|APPD\s*BY|DWN\s*BY|CHKO\.?\s*BY|"
    r"ISSUED\s*FOR|PLAN\s*HISTORY|SUBMITTED\s*TO|APPROVED\s*BY|"
    r"SHEET(\s*NO)?|SCALE|PROJECT\s*NO|PLAN\s*NO|DWG\s*NO|DRAWING\s*NO|"
    r"DEPT|DEPARTMENT|DSME"
    r")\b"
)

ALLCAPS_JUNK_WHITELIST = {"MAIN", "AUX", "E/R", "ICCP", "MGPS", "V/V", "V/D"}
VOWELS = set("AEIOU")

def extract_with_regex(text, patterns):
    results = {}
    for key, pattern in patterns.items():
        m = re.search(pattern, text or "", re.IGNORECASE)
        results[key] = m.group(1).strip() if m else "Unknown"
    return results

def normalize_text(text):
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("utf-8", errors="ignore")
    text = text.replace("\x00", " ")
    return re.sub(r"\s+", " ", text).strip()

def _alpha_ratio(s: str) -> float:
    if not s:
        return 0.0
    letters = sum(ch.isalpha() for ch in s)
    return letters / max(len(s), 1)

def _has_vowel(token: str) -> bool:
    return any(ch in VOWELS for ch in token.upper())

def _looks_like_stamp_or_revision(line_upper: str) -> bool:
    if not line_upper:
        return False
    if STAMP_REGEX.search(line_upper):
        return True
    if "OWN CHKD APPD" in line_upper or "DATE REV" in line_upper:
        return True
    return False

def _strip_stamp_fragments(s: str) -> str:
    if not s:
        return ""
    s = normalize_text(s)
    s = re.sub(r"(?i)\b(CHKD|APPD|DWN|CHKO)\b\s*\.?\s*(BY)?\s*[:\-]?\s*[A-Z.\s]{0,15}", " ", s)
    s = re.sub(r"(?i)\b(OWN|REV\.?|DATE|DESCRIPTION)\b\s*[:\-]?\s*[A-Z0-9.\s]{0,15}", " ", s)
    s = re.sub(r"(?i)\b(ISSUED\s*FOR|PLAN\s*HISTORY|SUBMITTED\s*TO|APPROVED\s*BY)\b.*", " ", s)
    s = STAMP_REGEX.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip()

def _normalize_title_terms(s: str) -> str:
    if not s:
        return ""
    s = normalize_text(s)

    # ✅ Canonicalize vendor drawing prefix FIRST
    # "V-D OF", "V D OF", "V.D OF", "V / D OF" -> "V/D of"
    s = re.sub(r"(?i)\bV\s*[-./]?\s*D\s*(?:OF)?\b", "V/D of", s)
    s = re.sub(r"(?i)\bV\s*/\s*D\b\s*OF\b", "V/D of", s)

    # Expand common abbreviations safely
    s = re.sub(r"(?i)\bPROV\.?\b", "Provision", s)
    s = re.sub(r"(?i)\bREF(?:ER)?\.?\b", "Refrigerating", s)  # REF / REFER -> Refrigerating
    s = re.sub(r"(?i)\bAIR[\-\s]?CON\b", "Air-Con", s)

    # ARR'T / ARR T / ARRANGMENT -> ARRANGEMENT
    s = re.sub(r"(?i)\bARR['’]?\s*T\b", "ARRANGEMENT", s)
    s = re.sub(r"(?i)\bARR\s*T\b", "ARRANGEMENT", s)
    s = re.sub(r"(?i)\bARRANGMENT\b", "ARRANGEMENT", s)

    # E-R / E R / E/R -> E/R
    s = re.sub(r"(?i)\bE\s*-\s*R\b", "E/R", s)
    s = re.sub(r"(?i)\bE\s+R\b", "E/R", s)
    s = re.sub(r"(?i)\bE\s*/\s*R\b", "E/R", s)

    # WORK SHOP -> WORKSHOP
    s = re.sub(r"(?i)\bWORK\s+SHOP\b", "WORKSHOP", s)

    # & -> AND
    s = re.sub(r"\s*&\s*", " AND ", s)

    return re.sub(r"\s+", " ", s).strip()

def _slice_from_first_anchor(s: str) -> str:
    if not s:
        return ""
    s = _normalize_title_terms(s)

    # ✅ If we already have "V/D of", never slice away its prefix
    if s.upper().startswith("V/D OF"):
        return s

    up = s.upper()
    best_idx = None
    for w in ANCHOR_WORDS:
        idx = up.find(w)
        if idx != -1:
            best_idx = idx if best_idx is None else min(best_idx, idx)

    if best_idx is None:
        return s

    head = s[:best_idx]
    if _alpha_ratio(head) < 0.40 or len(head) > 12:
        return s[best_idx:].strip()
    return s

def _drop_garbage_tokens(s: str) -> str:
    if not s:
        return ""
    tokens = []
    for tok in s.split():
        t = tok.strip()
        tu = t.upper()

        if tu in {"AND", "OF", "FOR", "IN"}:
            tokens.append(tu.lower() if tu in {"OF", "FOR", "IN"} else "and")
            continue

        if tu in {"E/R", "E-R"}:
            tokens.append("E/R")
            continue

        if tu == "V/D":
            tokens.append("V/D")
            continue

        if re.fullmatch(r"[\W_]{2,}", t):
            continue

        if len(t) <= 2 and tu not in {"ER"}:
            continue

        if re.search(r"[\"',`]", t) and _alpha_ratio(t) < 0.70:
            continue

        if t.isupper() and 3 <= len(t) <= 4 and tu not in ALLCAPS_JUNK_WHITELIST:
            # drop OLT/COL/DOL etc.
            if tu not in {"PLAN", "ROOM"}:
                continue

        if len(t) >= 3 and t.isalpha() and (not _has_vowel(t)) and tu not in ALLCAPS_JUNK_WHITELIST:
            continue

        if _alpha_ratio(t) < 0.25 and not re.search(r"[A-Za-z]{3,}", t):
            continue

        if re.search(r"[:;]{1,}", t) and _alpha_ratio(t) < 0.45:
            continue

        tokens.append(t)

    return re.sub(r"\s+", " ", " ".join(tokens)).strip()

def _is_meaningful_title(s: str) -> bool:
    if not s:
        return False
    up = s.upper()
    words = [w for w in s.split() if len(w) >= 3]
    if len(words) < 3 and not up.startswith("V/D OF"):
        return False
    if up.startswith("V/D OF"):
        return True
    if any(k in up for k in SYSTEM_KEYWORDS):
        return True
    if any(x in up for x in ["PLAN", "LIST", "ARRANGEMENT", "INSTALLATION", "DOOR", "INSULATION", "CRANE", "BEAM", "SYSTEM", "PLANT"]):
        return True
    return False

def clean_manual_name(name: str) -> str:
    if not name:
        return ""

    name = _normalize_title_terms(name)

    while True:
        prev = name

        name = re.sub(r"^[^A-Za-z0-9/&()]+|[^A-Za-z0-9/&()]+$", "", name).strip()
        name = _slice_from_first_anchor(name)
        name = _strip_stamp_fragments(name)
        name = _drop_garbage_tokens(name)
        name = _normalize_title_terms(name)

        # ✅ prefix removal (IMPORTANT: DO NOT REMOVE V/D)
        prefixes = [
            r"Manual for", r"Instruction for", r"Technical Manual for",
            r"Title", r"Ref", r"Technical", r"for",
            r"Instruction Manual -", r"Final & Instruction Manual for",
            r"Final Drawings", r"Instruction Manual", r"Technical Specification of",
            r"TITLE:", r"TITLE",
            r"DRAWING OF", r"DRAW I NG OF", r"VENDOR DWG OF", r"VENDOR DWG",
            r"DWG OF", r"VENDOR DRAWING OF", r"VENDOR DRAWING",
            # ❌ REMOVED: r"V\/D", r"V-D", r"V\sD", r"V\.D"
        ]
        pattern = r"^(?:" + "|".join(prefixes) + r")[\s:\-._]+"
        name = re.sub(pattern, "", name, flags=re.IGNORECASE).strip()

        if name == prev:
            break

    name = re.sub(r"\s*\(?\bT\s?E\s?L\b:?.*", "", name, flags=re.IGNORECASE).strip()
    name = re.sub(r"\s*\(?\b\d+(?:V|PH|HZ|KW)\b.*", "", name, flags=re.IGNORECASE).strip()

    name = re.sub(r"^[^A-Za-z0-9/&()]+|[^A-Za-z0-9/&()]+$", "", name).strip()
    name = re.sub(r"\s+", " ", name).strip()

    alnum = re.sub(r"[^a-zA-Z0-9]", "", name)
    if len(alnum) < 4 or name.upper() in {"TITLE", "MANUAL", "REF", "PROJECT", "DWG", "PLAN"}:
        return ""

    # Smart casing (keeps V/D, E/R, ICCP)
    def smart_title(s: str) -> str:
        out = []
        for w in s.split():
            wu = w.upper()
            if wu in {"E/R", "ICCP", "MGPS", "M.G.P.S", "V/V", "V/D"}:
                out.append(wu)
            elif w.isupper() and len(w) <= 4:
                out.append(w)
            else:
                out.append(w.capitalize())
        # Fix "V/D Of" to "V/D of"
        res = " ".join(out)
        res = re.sub(r"(?i)^V/D\s+Of\b", "V/D of", res)
        return res

    return smart_title(name)

def identify_manual_name(text, filename, folder_path="", metadata=None):
    text = text or ""
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    lines = lines[:900]

    filename_only = Path(filename).stem if filename else ""
    folder_name = Path(folder_path).name if folder_path else ""

    # Prefer filename for M(A)-xx, M(V)-xx, A(V)-xx drawings
    is_drawing_prefix = bool(re.match(r"(?i)^[A-Z]+\([A-Z]\)-\d+", filename_only or ""))
    fn_clean = re.sub(r"(?i)^[A-Z]+\([A-Z]\)-\d+\s*", "", filename_only).strip()
    fn_clean = fn_clean.replace("_", " ").replace("-", " ")
    fn_clean = re.sub(r"\s+", " ", fn_clean).strip()

    if is_drawing_prefix:
        words = [w for w in fn_clean.split() if len(w) >= 2]
        if len(words) >= 2:
            cleaned = clean_manual_name(fn_clean)
            if cleaned:
                return cleaned

    # ✅ Special: Vendor Drawing filenames A(V)-xx ... V-D OF ...
    if re.search(r"(?i)\bV\s*[-./]?\s*D\s*OF\b", fn_clean):
        forced = _normalize_title_terms(fn_clean)
        cleaned = clean_manual_name(forced)
        if cleaned:
            return cleaned

    # TITLE block extraction
    for i, line in enumerate(lines[:320]):
        m = re.search(r"\bTITLE\b\s*[:\-]?", line, flags=re.IGNORECASE)
        if not m:
            continue
        chunk = [line[m.start():]]
        for j in range(1, 35):
            if i + j >= len(lines):
                break
            nxt = lines[i + j].strip()
            nxt_up = nxt.upper()
            if not nxt:
                continue
            if _looks_like_stamp_or_revision(nxt_up):
                break
            if any(k in nxt_up for k in ["PROJECT NO", "PLAN NO", "DWG NO", "DRAWING NO", "SHEET", "SCALE", "DEPT", "DSME"]):
                break
            if _alpha_ratio(nxt) < 0.20 and not re.search(r"[A-Za-z]{3,}", nxt):
                continue
            chunk.append(nxt)
            joined = " ".join(chunk)
            if len(joined) > 40 and any(x in joined.upper() for x in ["V/D", "E/R", "PLAN", "LIST", "ARRANGEMENT", "SYSTEM", "PLANT"]):
                break

        cleaned = clean_manual_name(" ".join(chunk))
        if cleaned and _is_meaningful_title(cleaned):
            return cleaned

    # merged OCR: "... TITLE ... <title>"
    for line in lines[:420]:
        up = line.upper()
        if "TITLE" in up and any(k in up for k in ["V-D", "V/D", "SPARE", "ARR", "NAME", "VALVE", "EQUIPMENT", "SYSTEM", "PLANT"]):
            candidate = line[up.find("TITLE"):]
            cleaned = clean_manual_name(candidate)
            if cleaned and _is_meaningful_title(cleaned):
                return cleaned

    # keyword scan
    keywords = sorted(set(SYSTEM_KEYWORDS), key=len, reverse=True)
    for i, line in enumerate(lines[:550]):
        up = line.upper()
        if _looks_like_stamp_or_revision(up):
            continue
        if _alpha_ratio(line) < 0.18 and not re.search(r"[A-Za-z]{3,}", line):
            continue
        for kw in keywords:
            if kw in up and len(line) < 260:
                candidate = line
                for j in range(1, 12):
                    if i + j >= len(lines):
                        break
                    nxt = lines[i + j].strip()
                    nxt_up = nxt.upper()
                    if _looks_like_stamp_or_revision(nxt_up):
                        break
                    if _alpha_ratio(nxt) < 0.18 and not re.search(r"[A-Za-z]{3,}", nxt):
                        continue
                    candidate = f"{candidate} {nxt}"
                    if len(candidate) > 360:
                        break
                cleaned = clean_manual_name(candidate)
                if cleaned and _is_meaningful_title(cleaned):
                    return cleaned

    # label patterns
    for label in ["Document Title", "Engine Type", "Ship No", "Vessel Name"]:
        pat = LABEL_PATTERNS[label]
        found = re.search(pat, text, re.IGNORECASE)
        if found:
            cleaned = clean_manual_name(found.group(1).strip())
            if cleaned and _is_meaningful_title(cleaned):
                return cleaned

    # pdf metadata
    if metadata and isinstance(metadata, dict) and metadata.get("/Title"):
        title = str(metadata["/Title"]).strip()
        if len(title) > 5 and title.lower() not in {"untitled", "none", "1"}:
            cleaned = clean_manual_name(title)
            if cleaned and _is_meaningful_title(cleaned):
                return cleaned

    combo = f"{folder_name} - {fn_clean}".strip(" -")
    cleaned = clean_manual_name(combo)
    return cleaned if cleaned else ""

def classify_doc_type(text, filename, folder_path=""):
    combined = (normalize_text(text) + " " + normalize_text(filename) + " " + normalize_text(Path(folder_path).name)).lower()

    if any(k in combined for k in ["manual", "instruction", "handbook", "guide", "tmm", "operation manual", "maintenance manual"]):
        return "Machinery/System Manual"

    if any(k in combined for k in ["capacity plan", "tank capacity", "tank table", "sounding table", "deadweight scale"]):
        return "Capacity Plan / Datasheet"

    if any(k in combined for k in ["certificate", "report", "test result", "approval", "inclining experiment", "sea trial", "test record"]):
        return "Certificate / Report"

    if any(k in combined for k in ["drawing", "diagram", "schematic", "blueprint", "arr't", "arrangment", "arrangement", "plan", "details of", "v/d", "v/dwg", "vendor drawing"]):
        if "manual" not in combined:
            return "Drawing"

    if any(k in combined for k in [
        "equipment list", "v/v list", "valve list", "flowmeter", "plate type cooler", "centrifugal pump",
        "m.g.p.s", "prevention system", "name and caution plate",
        "door plan", "insulation plan", "installation of", "lifting beam", "crane",
        "air-con system", "provision refrigerating plant", "vendor drawing", "v/d of"
    ]):
        if "manual" not in combined:
            return "Drawing"

    if any(k in combined for k in ["structure", "construction", "body plan", "shell expansion", "deck and stringer", "bulkhead", "frames", "coamings", "fore body"]):
        return "Drawing"

    return "Unknown"