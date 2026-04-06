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


CATEGORY_RULES: Iterable[Tuple[str, Iterable[str]]] = [
    ("23", ("auto wash cloth", "wash cloth")),
    ("18", ("wash", "washing solution", "blanket wash", "roller wash", "uv wash")),
    ("19", ("fount", "fountain solution", "dampening solution")),
    ("20", ("plate cleaner", "plate care", "plate gum", "ctp cleaner")),
    ("21", ("roller care", "roller paste", "roller conditioner")),
    ("22", ("blanket care", "blanket reviver", "blanket maintenance")),
    ("01", ("rubber blanket", "printing blanket", "compressible blanket")),
    ("02", ("metalback blanket", "metal backed blanket")),
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


TYPE_KEYWORDS = (
    "wash",
    "blanket",
    "solution",
    "film",
    "foil",
    "paper",
    "matrix",
    "rule",
    "string",
    "rubber",
    "plate",
    "powder",
    "sponge",
    "hose",
    "tape",
)


def analyze_excel(file_obj) -> pd.DataFrame:
    df = pd.read_excel(file_obj)
    missing_columns = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing_columns:
        raise ValueError(
            "The Excel file must contain these columns exactly: "
            + ", ".join(REQUIRED_COLUMNS)
        )

    working_df = df[REQUIRED_COLUMNS].copy()
    working_df = working_df.fillna("")

    working_df["Brand"] = working_df["Item Name"].apply(extract_brand)
    working_df["Size"] = working_df.apply(extract_size, axis=1)
    working_df["Type"] = working_df.apply(extract_type, axis=1)
    working_df["Category"] = working_df.apply(extract_category, axis=1)

    return working_df


def extract_brand(item_name: str) -> str:
    text = normalize_spaces(item_name)
    if not text:
        return ""

    tokens = text.split()
    if len(tokens) >= 2 and tokens[1].lower() in {"tech", "teck"}:
        return f"{tokens[0]} {tokens[1]}"
    return tokens[0]


def extract_size(row: pd.Series) -> str:
    item_name = str(row["Item Name"])
    product_format = str(row["Product Format"])
    description = str(row["Description"])
    combined = " | ".join([item_name, product_format, description])

    patterns = [
        r"\b\d+(?:\.\d+)?\s?(?:ml|l|ltr|litre|liter)\b",
        r"\b\d+\s?(?:kg|g|gsm|mic|micron|mm|cm|m)\b",
        r"\b\d+\s?x\s?\d+(?:\s?x\s?\d+)?\s?(?:mm|cm|m|inch|in)\b",
        r"\b\d+\s?(?:tr|pcs|sheets|rolls)\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, combined, flags=re.IGNORECASE)
        if match:
            return normalize_spaces(match.group(0))

    return normalize_spaces(product_format)


def extract_type(row: pd.Series) -> str:
    item_name = normalize_spaces(str(row["Item Name"]))
    description = normalize_spaces(str(row["Description"]))
    product_format = normalize_spaces(str(row["Product Format"]))
    haystack = f"{item_name} {description}".lower()

    for keyword in TYPE_KEYWORDS:
        if keyword in haystack:
            if keyword == "wash":
                if "uv" in haystack:
                    return "UV Wash"
                if "auto" in haystack:
                    return "Auto Wash"
                return "Wash"
            if keyword == "solution":
                if "fount" in haystack or "fountain" in haystack:
                    return "Fountain Solution"
                return "Solution"
            return keyword.title()

    return product_format or item_name


def extract_category(row: pd.Series) -> str:
    haystack = " ".join(
        [
            normalize_spaces(str(row["Item Name"])),
            normalize_spaces(str(row["Description"])),
            normalize_spaces(str(row["Product Format"])),
        ]
    ).lower()

    for category_number, keywords in CATEGORY_RULES:
        if any(keyword in haystack for keyword in keywords):
            return format_category(category_number)

    item_name = str(row["Item Name"]).lower()
    if "wash" in item_name:
        return format_category("18")
    if "fount" in item_name or "fountain" in item_name:
        return format_category("19")

    return "Unmapped - Review Needed"


def format_category(category_number: str) -> str:
    return f"{category_number} - {CATEGORY_MAP[category_number]}"


def normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", str(value)).strip()
