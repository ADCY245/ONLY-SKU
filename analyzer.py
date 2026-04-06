import re
from typing import Dict, Iterable, Tuple

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
    "day": "Day",
    "phoenix": "Phoenix",
    "vulcan": "Vulcan",
    "contitech": "ContiTech",
    "kinyo": "Kinyo",
    "bottcher": "Bottcher",
}

NON_BRAND_PREFIXES = {"pl", "d", "alub", "exsq"}

CATEGORY_RULES: Iterable[Tuple[str, Iterable[str]]] = [
    ("23", ("auto wash cloth", "wash cloth")),
    ("19", ("fount", "fountain solution", "dampening solution")),
    ("20", ("plate cleaner", "plate care", "plate gum", "ctp cleaner")),
    ("21", ("roller care", "roller paste", "roller conditioner")),
    ("22", ("blanket care", "blanket reviver", "blanket maintenance")),
    ("18", ("washing solution", "blanket wash", "roller wash", "uv wash", "wash")),
    ("02", ("metalback blanket", "metal backed blanket", "alub")),
    ("01", ("rubber blanket", "printing blanket", "compressible blanket", "sheet-fed blanket")),
    ("03", ("underlay blanket",)),
    ("04", ("blanket barring",)),
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
    working_df["Size"] = working_df.apply(extract_size, axis=1)
    working_df["Product Format"] = working_df.apply(normalize_product_format, axis=1)
    working_df["Type"] = working_df.apply(extract_type, axis=1)
    working_df["Category"] = working_df.apply(extract_category, axis=1)
    return working_df


def extract_brand(item_name: str) -> str:
    text = normalize_spaces(item_name).lower()
    if not text:
        return ""

    for brand, display_name in KNOWN_BRANDS.items():
        if re.search(rf"\b{re.escape(brand)}\b", text):
            return display_name

    cleaned_tokens = [
        token
        for token in re.split(r"[\s|/-]+", text)
        if token and token not in NON_BRAND_PREFIXES and not re.fullmatch(r"\d+(?:\.\d+)?", token)
    ]
    return cleaned_tokens[0].title() if cleaned_tokens else ""


def extract_size(row: pd.Series) -> str:
    combined = " | ".join(
        [str(row["Item Name"]), str(row["Product Format"]), str(row["Description"])]
    )
    patterns = [
        r"\b\d+(?:\.\d+)?\s?(?:mm|cm|m|inch|in)\s?x\s?\d+(?:\.\d+)?\s?(?:mm|cm|m|inch|in)(?:\s?x\s?\d+(?:\.\d+)?\s?(?:mm|cm|m|inch|in))?\b",
        r"\b\d+(?:\.\d+)?\s?(?:ml|l|ltr|litre|liter)\b",
        r"\b\d+(?:\.\d+)?\s?(?:mm|cm|m|mic|micron)\b",
        r"\b\d+(?:\.\d+)?\s?(?:kg|g|gsm)\b",
        r"\b\d+\s?(?:tr|pcs|sheets|rolls)\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, combined, flags=re.IGNORECASE)
        if match:
            return normalize_spaces(match.group(0))
    return normalize_spaces(str(row["Product Format"]))


def normalize_product_format(row: pd.Series) -> str:
    product_format = normalize_spaces(str(row["Product Format"]))
    if product_format and not is_thickness_only(product_format):
        return product_format
    if is_blanket_product(row):
        return "Rubber Blanket - Roll Format"
    return product_format


def extract_type(row: pd.Series) -> str:
    haystack = build_haystack(row)
    size = normalize_spaces(str(row["Size"]))

    if is_blanket_product(row):
        return "Rubber Blanket"
    if is_liter_product(size):
        if "fount" in haystack or "fountain" in haystack:
            return "Fountain Solution"
        if "wash" in haystack:
            return "Wash"
        if "plate" in haystack:
            return "Plate Care Product"
        if "roller" in haystack:
            return "Roller Care Product"
        return "Chemical / Maintenance Product"
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
    if "sponge" in haystack:
        return "Sponge"
    return normalize_spaces(str(row["Product Format"])) or normalize_spaces(str(row["Item Name"]))


def extract_category(row: pd.Series) -> str:
    haystack = build_haystack(row)
    size = normalize_spaces(str(row["Size"]))

    if is_blanket_product(row):
        if any(keyword in haystack for keyword in ("metalback", "metal backed", "alub")):
            return format_category("02")
        return format_category("01")

    if is_liter_product(size):
        if "fount" in haystack or "fountain" in haystack:
            return format_category("19")
        if "plate" in haystack:
            return format_category("20")
        if "roller" in haystack:
            return format_category("21")
        if "blanket" in haystack and any(
            keyword in haystack for keyword in ("care", "reviver", "maint", "maintenance")
        ):
            return format_category("22")
        return format_category("18")

    for category_number, keywords in CATEGORY_RULES:
        if any(keyword in haystack for keyword in keywords):
            return format_category(category_number)
    return "Unmapped - Review Needed"


def is_blanket_product(row: pd.Series) -> bool:
    haystack = build_haystack(row)
    size = normalize_spaces(str(row["Size"]))
    has_thickness = bool(re.search(r"\b\d+(?:\.\d+)?\s?mm\b", size, flags=re.IGNORECASE))

    if any(
        keyword in haystack
        for keyword in ("blanket", "uv black", "webline", "topaz", "privilege", "advantage plus", "magnum")
    ):
        return True
    if has_thickness and not is_liter_product(size):
        return True
    return False


def is_liter_product(size: str) -> bool:
    return bool(re.search(r"\b\d+(?:\.\d+)?\s?(?:ml|l|ltr|litre|liter)\b", size, flags=re.IGNORECASE))


def is_thickness_only(product_format: str) -> bool:
    return bool(re.fullmatch(r"\d+(?:\.\d+)?\s?mm", product_format, flags=re.IGNORECASE))


def build_haystack(row: pd.Series) -> str:
    return " ".join(
        [
            normalize_spaces(str(row["Item Name"])),
            normalize_spaces(str(row["Description"])),
            normalize_spaces(str(row["Product Format"])),
        ]
    ).lower()


def format_category(category_number: str) -> str:
    return f"{category_number} - {CATEGORY_MAP[category_number]}"


def normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", str(value)).strip()
