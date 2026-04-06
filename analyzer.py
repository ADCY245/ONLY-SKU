import re
from typing import Dict, Iterable, List, Tuple

import pandas as pd


REQUIRED_COLUMNS = ["Item Name", "Description", "Product Format"]

CATEGORY_MAP: Dict[str, str] = {
    "01": "Rubber Blankets",
    "02": "Metalback Blankets",
    "03": "Underlay Blanket",
    "04": "Blanket Barring",
    "05": "Calibrated Underpacking Paper",
    "06": "Calibrated Underpacking Film",
    "07": "Creasing Matrix",
    "08": "Cutting Rules",
    "09": "Creasing Rules",
    "10": "Litho Perforation Rules",
    "11": "Cutting String",
    "12": "Ejection Rubber",
    "13": "Strip Plate",
    "14": "Anti Marking Film",
    "15": "Ink Duct Foil",
    "16": "Productive Foil",
    "17": "Presspahn Sheets",
    "18": "Washing Solutions",
    "19": "Fountain Solutions",
    "20": "Plate Care Products",
    "21": "Roller Care Products",
    "22": "Blanket Maintenance Products",
    "23": "Auto Wash Cloth",
    "24": "ICP Paper",
    "25": "Spray Powder",
    "26": "Sponges",
    "27": "Dampening Hose",
    "28": "Tesamol Tape",
}


KNOWN_BRANDS = {
    "sava": "Sava",
    "image": "Image",
    "mtech": "MTech",
    "mteck": "MTeck",
    "m3z": "Marks3.Zet",
    "marks3.zet": "Marks3.Zet",
    "mpack": "MPack",
    "polipack": "Polipack",
    "b4p": "B4P",
    "sigma": "Sigma",
    "star": "Star",
    "fujikura": "Fujikura",
    "thompson": "Thompson",
    "hs boyd": "HS Boyd",
    "day": "Day",
    "phoenix": "Phoenix",
    "vulcan": "Vulcan",
    "contitech": "ContiTech",
    "kinyo": "Kinyo",
    "bottcher": "Bottcher",
}

NON_BRAND_PREFIXES = {"pl", "d", "exsq"}
CODES_TO_REMOVE = {"alub", "stlb", "exsq"}
MACHINE_NAMES_TO_REMOVE = ("kba rapida", "kba rabida")

CATEGORY_RULES: Iterable[Tuple[str, Iterable[str]]] = [
    ("23", ("auto wash cloth", "wash cloth")),
    ("19", ("fount", "fountain solution", "dampening solution")),
    ("20", ("plate cleaner", "plate care", "plate gum", "ctp cleaner")),
    ("21", ("roller care", "roller paste", "roller conditioner")),
    ("22", ("blanket care", "blanket reviver", "blanket maintenance")),
    ("18", ("washing solution", "blanket wash", "roller wash", "uv wash", "wash gp", "wash hsw", "wash auto", "wash")),
    ("04", ("barring", "b4p")),
    ("02", ("metalback blanket", "metal backed blanket", "alub", "stlb")),
    ("01", ("rubber blanket", "printing blanket", "compressible blanket", "sheet-fed blanket")),
    ("03", ("underlay blanket", "underpacking blanket")),
    ("05", ("underpacking paper", "calibrated underpacking paper")),
    ("06", ("underpacking film", "calibrated underpacking film")),
    ("07", ("creasing matrix", "matrix")),
    ("08", ("cutting rule", "cutting rules")),
    ("09", ("creasing rule", "creasing rules")),
    ("10", ("perforation rule", "litho perforation")),
    ("11", ("cutting string",)),
    ("12", ("ejection rubber",)),
    ("13", ("strip plate",)),
    ("14", ("anti marking film", "anti-marking film")),
    ("15", ("ink duct foil",)),
    ("16", ("productive foil",)),
    ("17", ("presspahn", "presspahn sheets")),
    ("24", ("icp paper",)),
    ("25", ("spray powder",)),
    ("26", ("sponge", "sponges")),
    ("27", ("dampening hose",)),
    ("28", ("tesamol tape", "tesamol")),
]


def analyze_excel(file_obj) -> pd.DataFrame:
    df = pd.read_excel(file_obj)
    missing_columns = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing_columns:
        raise ValueError(
            "The Excel file must contain these columns exactly: "
            + ", ".join(REQUIRED_COLUMNS)
        )

    working_df = df[REQUIRED_COLUMNS].copy().fillna("")
    working_df["Brand"] = working_df["Item Name"].apply(extract_brand)
    working_df["Product Name"] = working_df.apply(extract_product_name, axis=1)
    working_df["Size"] = working_df.apply(extract_size, axis=1)
    working_df["Product Format"] = working_df.apply(normalize_product_format, axis=1)
    working_df["Type"] = working_df.apply(extract_type, axis=1)
    working_df["Category"] = working_df.apply(extract_category, axis=1)
    return working_df


def extract_brand(item_name: str) -> str:
    text = normalize_spaces(item_name).lower()
    if not text:
        return "Unspecified"

    for brand, display_name in KNOWN_BRANDS.items():
        if re.search(rf"\b{re.escape(brand)}\b", text):
            return display_name
    return "Unspecified"


def extract_product_name(row: pd.Series) -> str:
    item_name = normalize_spaces(str(row["Item Name"]))
    if not item_name:
        return ""

    brand = str(row.get("Brand", ""))
    if brand in {"Thompson", "HS Boyd"}:
        return extract_thompson_hs_boyd_product_name(item_name, brand)

    text = item_name
    text = re.sub(r"^[A-Za-z]{1,3}\s*[|/-]\s*", "", text)
    for machine_name in MACHINE_NAMES_TO_REMOVE:
        text = re.sub(rf"\b{re.escape(machine_name)}\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(
        r"\b\d+(?:\.\d+)?\s?(?:mm|cm|m|mtr|meter|meters|inch|in)\s*x\s*\d+(?:\.\d+)?\s?(?:mm|cm|m|mtr|meter|meters|inch|in)(?:\s*x\s*\d+(?:\.\d+)?\s?(?:mm|cm|m|mtr|meter|meters|inch|in))?\b",
        " ",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(r"\b\d+(?:\.\d+)?\s?mm\s*x\s*\d+(?:\.\d+)?\s*pt\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?\s*pt\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b\d+(?:\.\d+)?\s?(?:ml|l|ltr|litre|liter|mm|cm|m|mtr|meter|meters|kg|g|gsm)\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b\d{4,}\b", " ", text)
    for code in CODES_TO_REMOVE:
        text = re.sub(rf"\b{re.escape(code)}\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\b(?:cfr|cf)\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bglossy\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"[()]", " ", text)
    text = re.sub(r"\bself\s+adhesive\b", "Self adhesive", text, flags=re.IGNORECASE)
    text = re.sub(r"\b\d+(?:\.\d+)?\s?(?:mic|micron)\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*[|/-]\s*", " ", text)
    text = re.sub(r"\s+", " ", text).strip(" -|/")

    mic_match = re.search(r"\b\d+(?:\.\d+)?\s?(?:mic|micron)\b", item_name, flags=re.IGNORECASE)
    if mic_match and "self adhesive" in text.lower():
        text = f"{text} - {normalize_spaces(mic_match.group(0))}"

    return normalize_spaces(text)


def extract_size(row: pd.Series) -> str:
    item_name = normalize_spaces(str(row["Item Name"]))
    brand = str(row.get("Brand", ""))
    if brand in {"Thompson", "HS Boyd"}:
        return extract_thompson_hs_boyd_size(item_name)

    combined = " | ".join(
        [str(row["Item Name"]), str(row["Product Format"]), str(row["Description"])]
    )
    patterns = [
        r"\b\d+(?:\.\d+)?\s?(?:mm|cm|m|mtr|meter|meters|inch|in)\s?x\s?\d+(?:\.\d+)?\s?(?:mm|cm|m|mtr|meter|meters|inch|in)(?:\s?x\s?\d+(?:\.\d+)?\s?(?:mm|cm|m|mtr|meter|meters|inch|in))?\b",
        r"\b\d+(?:\.\d+)?\s?mm\s*x\s*\d+(?:\.\d+)?\s*pt\b",
        r"\b\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?\s*pt\b",
        r"\b\d+(?:\.\d+)?\s?(?:ml|l|ltr|litre|liter)\b",
        r"\b\d+(?:\.\d+)?\s?(?:mm|cm|m|mtr|meter|meters|mic|micron)\b",
        r"\b\d+(?:\.\d+)?\s?(?:kg|g|gsm)\b",
        r"\b\d+\s?(?:tr|pcs|sheets|rolls)\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, combined, flags=re.IGNORECASE)
        if match:
            return normalize_spaces(match.group(0))
    return normalize_spaces(str(row["Product Format"]))


def normalize_product_format(row: pd.Series) -> str:
    existing_format = normalize_spaces(str(row["Product Format"]))
    classified_type = classify_type_label(row)
    if classified_type in {
        "Rubber Blanket - Cut Format",
        "Rubber Blanket - Roll Format",
        "Rubber Blanket - Bar Cut Format",
        "Underpacking - Cut Format",
        "Underpacking - Roll Format",
        "Barring Pieces",
        "Sponge Pieces",
        "Creasing Matrix",
        "Cutting Rule",
        "Creasing Rule",
        "Litho Perforation Rule",
    }:
        return classified_type
    return existing_format


def extract_type(row: pd.Series) -> str:
    return classify_type_label(row)


def extract_category(row: pd.Series) -> str:
    haystack = build_haystack(row)
    size = normalize_spaces(str(row["Size"]))
    item_name = normalize_spaces(str(row["Item Name"])).lower()
    brand = str(row.get("Brand", ""))

    if is_wash_product(row):
        return format_category("18")
    if is_fountain_product(row):
        return format_category("19")
    if is_plate_care_product(row):
        return format_category("20")
    if is_roller_care_product(row):
        return format_category("21")
    if is_blanket_maintenance_product(row):
        return format_category("22")
    if is_barring_piece_product(row):
        return format_category("04")
    if is_matrix_product(row):
        return format_category("07")
    if is_rule_product(row):
        if is_perforation_rule_product(row):
            return format_category("10")
        if is_creasing_rule_product(row):
            return format_category("09")
        return format_category("08")
    if brand == "Marks3.Zet":
        return format_category("05")
    if brand == "Polipack":
        return format_category("06")
    if is_anti_marking_product(row):
        return format_category("06")
    if is_underpacking_product(row):
        if "film" in haystack:
            return format_category("06")
        if "paper" in haystack:
            return format_category("05")
        return format_category("03")
    if is_blanket_product(row):
        if "mbb" in haystack:
            return format_category("02")
        return format_category("01")
    if is_sponge_product(row):
        return format_category("26")
    if is_liter_product(size):
        return format_category("18")

    for category_number, keywords in CATEGORY_RULES:
        if any(keyword in item_name or keyword in haystack for keyword in keywords):
            return format_category(category_number)
    return "Unmapped - Review Needed"


def classify_type_label(row: pd.Series) -> str:
    haystack = build_haystack(row)
    size = normalize_spaces(str(row["Size"]))
    brand = str(row.get("Brand", ""))

    if is_wash_product(row):
        return "Washing Solution"
    if is_fountain_product(row):
        return "Fountain Solution"
    if is_plate_care_product(row):
        return "Plate Care Product"
    if is_roller_care_product(row):
        return "Roller Care Product"
    if is_blanket_maintenance_product(row):
        return "Blanket Maintenance Product"
    if is_barring_piece_product(row):
        return "Barring Pieces"
    if is_sponge_product(row):
        return "Sponge Pieces"
    if is_matrix_product(row):
        return "Creasing Matrix"
    if is_rule_product(row):
        if is_perforation_rule_product(row):
            return "Litho Perforation Rule"
        if is_creasing_rule_product(row):
            return "Creasing Rule"
        return "Cutting Rule"
    if brand == "Marks3.Zet":
        return "Underpacking - Paper"
    if brand == "Polipack":
        return "Underpacking - Film"
    if is_anti_marking_product(row):
        return "Underpacking - Film"
    if is_underpacking_product(row):
        if is_cut_dimensions(size):
            return "Underpacking - Cut Format"
        if is_roll_dimensions(size):
            return "Underpacking - Roll Format"
        return "Underpacking"
    if is_blanket_product(row):
        if has_bar_cut_code(haystack):
            return "Rubber Blanket - Bar Cut Format"
        if is_cut_dimensions(size):
            return "Rubber Blanket - Cut Format"
        if is_roll_dimensions(size) or is_thickness_only(size):
            return "Rubber Blanket - Roll Format"
        return "Rubber Blanket"
    if "film" in haystack:
        return "Film"
    if "foil" in haystack:
        return "Foil"
    if "paper" in haystack:
        return "Paper"
    if "matrix" in haystack:
        return "Creasing Matrix"
    if "rule" in haystack:
        return "Rule"
    if "tape" in haystack:
        return "Tape"
    if "hose" in haystack:
        return "Hose"
    if "powder" in haystack:
        return "Powder"
    return normalize_spaces(str(row["Product Format"])) or normalize_spaces(str(row["Item Name"]))


def is_blanket_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    size = normalize_spaces(str(row["Size"]))
    if is_liter_product(size):
        return False
    if is_underpacking_product(row):
        return False
    if is_barring_piece_product(row):
        return False
    if is_matrix_product(row):
        return False
    if is_rule_product(row):
        return False
    if is_anti_marking_product(row):
        return False
    if any(
        keyword in haystack
        for keyword in ("blanket", "uv black", "webline", "topaz", "privilege", "advantage plus", "magnum", "print master", "web master", "mbb")
    ):
        return True
    if re.search(r"\b\d+(?:\.\d+)?\s?mm\b", size, flags=re.IGNORECASE):
        return True
    return False


def is_underpacking_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    return any(keyword in haystack for keyword in ("mz", "underpacking", "underlay"))


def is_matrix_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    return "matrix" in haystack


def is_rule_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    brand = str(row.get("Brand", ""))
    size = normalize_spaces(str(row.get("Size", "")))
    if "rule" in haystack:
        return True
    if brand in {"Thompson", "HS Boyd"} and any(keyword in haystack for keyword in ("perforation", "perf", "tpi", "paper", "card", "side", "centre", "center")):
        return True
    if brand in {"Sigma", "Star", "Fujikura"} and bool(re.search(r"\b\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?\s*pt\b", size, flags=re.IGNORECASE)):
        return True
    return False


def is_anti_marking_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    return "anti marking" in haystack or "anti-marking" in haystack


def is_creasing_rule_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    product_name = normalize_spaces(str(row.get("Product Name", ""))).lower()
    return any(keyword in haystack for keyword in ("creasing rule", "crease")) or "lcb" in product_name


def is_perforation_rule_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    return "perforation" in haystack or "perf" in haystack


def is_barring_piece_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    return "b4p" in haystack or "barring piece" in haystack or "barring pieces" in haystack


def is_sponge_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    return "sponge" in haystack


def is_wash_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    size = normalize_spaces(str(row["Size"]))
    return is_liter_product(size) and "wash" in haystack


def is_fountain_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    return "fount" in haystack or "fountain" in haystack


def is_plate_care_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    return any(keyword in haystack for keyword in ("plate care", "plate cleaner", "plate gum", "ctp cleaner"))


def is_roller_care_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    return any(keyword in haystack for keyword in ("roller care", "roller paste", "roller conditioner"))


def is_blanket_maintenance_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    return any(keyword in haystack for keyword in ("blanket care", "blanket reviver", "blanket maintenance"))


def is_liter_product(size: str) -> bool:
    return bool(re.search(r"\b\d+(?:\.\d+)?\s?(?:ml|l|ltr|litre|liter)\b", size, flags=re.IGNORECASE))


def is_thickness_only(value: str) -> bool:
    return bool(re.fullmatch(r"\d+(?:\.\d+)?\s?mm", value, flags=re.IGNORECASE))


def is_cut_dimensions(size: str) -> bool:
    units = extract_dimension_units(size)
    return len(units) == 3 and units[0] == "mm" and units[1] == "mm" and units[2] == "mm"


def is_roll_dimensions(size: str) -> bool:
    units = extract_dimension_units(size)
    return len(units) == 3 and units[0] == "mm" and units[1] in {"m", "mtr", "meter", "meters"} and units[2] == "mm"


def extract_dimension_units(size: str) -> List[str]:
    parts = re.split(r"\s*x\s*", size, flags=re.IGNORECASE)
    units: List[str] = []
    for part in parts:
        match = re.search(r"(mm|cm|mtr|meters|meter|m|inch|in)\b", part, flags=re.IGNORECASE)
        if match:
            units.append(match.group(1).lower())
    return units


def has_bar_cut_code(haystack: str) -> bool:
    return any(code in haystack for code in ("alub", "stlb"))


def extract_thompson_hs_boyd_product_name(item_name: str, brand: str) -> str:
    number_match = re.search(r"\b([A-Z]\d{2,}[A-Z0-9-]*)\b", item_name, flags=re.IGNORECASE)
    if number_match:
        return f"{brand} {number_match.group(1).upper()}"
    return brand


def extract_thompson_hs_boyd_size(item_name: str) -> str:
    parts: List[str] = []

    tpi_match = re.search(r"\b\d+(?:\.\d+)?\s*TPI\b", item_name, flags=re.IGNORECASE)
    if tpi_match:
        parts.append(normalize_spaces(tpi_match.group(0)).upper())

    side_match = re.search(r"\b(side|centre|center)\b", item_name, flags=re.IGNORECASE)
    if side_match:
        side_value = side_match.group(1).lower()
        if side_value == "center":
            side_value = "centre"
        parts.append(side_value.title())

    stock_match = re.search(r"\b(paper|card)\b", item_name, flags=re.IGNORECASE)
    if stock_match:
        parts.append(stock_match.group(1).title())

    return " - ".join(parts)


def build_haystack(row: pd.Series) -> str:
    return " ".join(
        [
            normalize_spaces(str(row["Item Name"])),
            normalize_spaces(str(row["Description"])),
            normalize_spaces(str(row["Product Format"])),
            normalize_spaces(str(row.get("Product Name", ""))),
        ]
    ).lower()


def format_category(category_number: str) -> str:
    return f"{category_number} - {CATEGORY_MAP[category_number]}"


def normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", str(value)).strip()
