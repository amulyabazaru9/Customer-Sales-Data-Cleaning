"""
generate_dataset.py
-------------------
Generates a realistic but *intentionally messy* customer sales dataset.

The goal is to simulate the kind of file a small business might hand to a
freelancer: exported from several systems, typed in by hand, full of
duplicates, inconsistent formats and missing values.

All people, emails and phone numbers are fictional. Emails use the reserved
example.com / example.org / example.net domains, and phone numbers use the
fictional 555 exchange.

Usage (from the project root):
    python generate_dataset.py
    python generate_dataset.py --rows 500 --seed 42
"""

from __future__ import annotations

import argparse
import csv
import random
import sys
from datetime import date, timedelta
from pathlib import Path

# ---------------------------------------------------------------------------
# Paths are built relative to this file, so the script works from any machine
# (no hard-coded C:\Users\... paths).
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DEFAULT_OUTPUT = DATA_DIR / "raw_sales_data.csv"

COLUMNS = [
    "order_id",
    "order_date",
    "customer_name",
    "customer_email",
    "phone",
    "product",
    "category",
    "quantity",
    "unit_price",
    "city",
    "state",
]

# Product catalogue: (product name, category, list price in USD)
PRODUCTS = [
    ("Wireless Mouse", "Accessories", 24.99),
    ("Mechanical Keyboard", "Accessories", 89.99),
    ("USB-C Hub", "Accessories", 39.99),
    ("Laptop Stand", "Accessories", 34.50),
    ("Noise Cancelling Headphones", "Audio", 199.00),
    ("Bluetooth Speaker", "Audio", 59.95),
    ("27-inch 4K Monitor", "Displays", 349.00),
    ("Webcam HD", "Video", 69.00),
    ("External SSD 1TB", "Storage", 119.99),
    ("Office Chair", "Furniture", 249.00),
    ("Standing Desk", "Furniture", 499.00),
    ("Laptop Pro 14", "Computers", 1299.00),
]

# (city, state) pairs used as customer locations
LOCATIONS = [
    ("New York", "NY"),
    ("Los Angeles", "CA"),
    ("Chicago", "IL"),
    ("Houston", "TX"),
    ("Phoenix", "AZ"),
    ("Philadelphia", "PA"),
    ("San Diego", "CA"),
    ("Dallas", "TX"),
    ("Seattle", "WA"),
    ("Miami", "FL"),
    ("Denver", "CO"),
    ("Boston", "MA"),
]

FIRST_NAMES = [
    "James", "Mary", "Robert", "Patricia", "John", "Jennifer", "Michael",
    "Linda", "David", "Elizabeth", "William", "Barbara", "Richard", "Susan",
    "Joseph", "Jessica", "Thomas", "Sarah", "Daniel", "Karen", "Aisha",
    "Carlos", "Mei", "Priya", "Omar", "Sofia", "Lucas", "Fatima", "Noah",
    "Emma", "Liam", "Olivia", "Ethan", "Amara", "Hiro", "Elena",
]
LAST_NAMES = [
    "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller",
    "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Wilson",
    "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Lee", "Patel",
    "Khan", "Nguyen", "Kim", "Silva", "Okafor", "Tanaka", "Rossi", "Clark",
]
EMAIL_DOMAINS = ["example.com", "example.org", "example.net"]

# Several date formats, as if the file was merged from different systems.
# Every format is unambiguous on its own, so the cleaner can parse it safely.
DATE_FORMATS = [
    "%Y-%m-%d",    # 2025-03-14
    "%d/%m/%Y",    # 14/03/2025
    "%m-%d-%Y",    # 03-14-2025
    "%B %d, %Y",   # March 14, 2025
    "%Y/%m/%d",    # 2025/03/14
    "%d-%b-%Y",    # 14-Mar-2025
]

START_DATE = date(2025, 1, 1)
END_DATE = date(2025, 12, 31)


# ---------------------------------------------------------------------------
# Helpers that create "messy" versions of clean values
# ---------------------------------------------------------------------------
def messy_case(text: str, rng: random.Random) -> str:
    """Return the text in a random (often wrong) capitalisation."""
    return rng.choice([text.upper(), text.lower(), text.swapcase(), text])


def messy_spaces(text: str, rng: random.Random) -> str:
    """Add leading/trailing spaces and double spaces between words."""
    text = text.replace(" ", rng.choice(["  ", " ", "   "]))
    return rng.choice([" ", "  ", ""]) + text + rng.choice([" ", "   ", "\t"])


def format_phone(digits: str, rng: random.Random) -> str:
    """Format a 10-digit phone number in one of several common styles."""
    a, b, c = digits[:3], digits[3:6], digits[6:]
    return rng.choice([
        f"({a}) {b}-{c}",
        f"{a}-{b}-{c}",
        f"{a}.{b}.{c}",
        f"{a}{b}{c}",
        f"+1 {a} {b} {c}",
        f"+1-{a}-{b}-{c}",
    ])


def build_customers(rng: random.Random, count: int = 120) -> list[dict]:
    """Create a pool of fictional customers who place repeat orders."""
    customers = []
    used_emails = set()
    while len(customers) < count:
        first, last = rng.choice(FIRST_NAMES), rng.choice(LAST_NAMES)
        email = f"{first}.{last}{rng.randint(1, 99)}@{rng.choice(EMAIL_DOMAINS)}".lower()
        if email in used_emails:
            continue
        used_emails.add(email)
        city, state = rng.choice(LOCATIONS)
        customers.append({
            "name": f"{first} {last}",
            "email": email,
            "phone": f"555{rng.randint(1000000, 9999999)}",
            "city": city,
            "state": state,
        })
    return customers


def random_date(rng: random.Random) -> date:
    span = (END_DATE - START_DATE).days
    return START_DATE + timedelta(days=rng.randint(0, span))


def make_clean_order(order_no: int, customers: list[dict], rng: random.Random) -> dict:
    """Create one *correct* order before any data-quality problems are added."""
    cust = rng.choice(customers)
    product, category, price = rng.choice(PRODUCTS)
    # Cheaper products tend to be bought in larger quantities
    max_qty = 8 if price < 100 else 3
    return {
        "order_id": f"ORD-{order_no}",
        "order_date": random_date(rng),
        "customer_name": cust["name"],
        "customer_email": cust["email"],
        "phone": cust["phone"],
        "product": product,
        "category": category,
        "quantity": rng.randint(1, max_qty),
        "unit_price": price,
        "city": cust["city"],
        "state": cust["state"],
    }


def add_quality_problems(row: dict, rng: random.Random) -> dict:
    """
    Inject realistic data-quality problems into a single row.
    Probabilities are tuned so most rows are still usable, which mirrors
    real business data: messy but recoverable.
    """
    r = dict(row)

    # 1. Inconsistent date formats (every row) + a few invalid / missing dates
    roll = rng.random()
    if roll < 0.015:
        r["order_date"] = ""
    elif roll < 0.03:
        r["order_date"] = rng.choice(["N/A", "unknown", "2025-13-45", "31/02/2025"])
    else:
        r["order_date"] = r["order_date"].strftime(rng.choice(DATE_FORMATS))

    # 2. Order IDs typed in lowercase or with spaces; a couple missing
    roll = rng.random()
    if roll < 0.01:
        r["order_id"] = ""
    elif roll < 0.10:
        r["order_id"] = messy_spaces(r["order_id"].lower(), rng)

    # 3. Customer names: bad capitalisation, extra spaces, some missing
    roll = rng.random()
    if roll < 0.04:
        r["customer_name"] = ""
    else:
        if rng.random() < 0.20:
            r["customer_name"] = messy_case(r["customer_name"], rng)
        if rng.random() < 0.15:
            r["customer_name"] = messy_spaces(r["customer_name"], rng)

    # 4. Emails: uppercase / spaces (fixable), malformed or missing (not fixable)
    roll = rng.random()
    if roll < 0.04:
        r["customer_email"] = ""
    elif roll < 0.08:
        email = r["customer_email"]
        r["customer_email"] = rng.choice([
            email.replace("@", ""),            # missing @
            email.replace("@", "@@"),          # double @
            email.split("@")[0] + "@",         # missing domain
            email.replace(".com", ".").replace(".org", ".").replace(".net", "."),
        ])
    elif roll < 0.20:
        r["customer_email"] = messy_spaces(r["customer_email"].upper(), rng)

    # 5. Phone numbers in many formats, some missing or too short
    roll = rng.random()
    if roll < 0.04:
        r["phone"] = ""
    elif roll < 0.06:
        r["phone"] = r["phone"][:rng.randint(5, 8)]  # truncated / invalid
    else:
        r["phone"] = format_phone(r["phone"], rng)

    # 6. Product names with wrong case / extra spaces
    if rng.random() < 0.15:
        r["product"] = messy_case(r["product"], rng)
    if rng.random() < 0.10:
        r["product"] = messy_spaces(r["product"], rng)
    if rng.random() < 0.10:
        r["category"] = messy_case(r["category"], rng)

    # 7. Quantities: decimals like "2.0" (fixable), zero/negative/text/missing (invalid)
    roll = rng.random()
    if roll < 0.015:
        r["quantity"] = ""
    elif roll < 0.035:
        r["quantity"] = rng.choice([0, -1, -3])
    elif roll < 0.045:
        r["quantity"] = rng.choice(["two", "N/A", "?"])
    elif roll < 0.10:
        r["quantity"] = f"{r['quantity']}.0"

    # 8. Prices: currency symbols (fixable), negative/zero/text/outlier/missing (invalid)
    roll = rng.random()
    price = r["unit_price"]
    if roll < 0.015:
        r["unit_price"] = ""
    elif roll < 0.035:
        r["unit_price"] = -price
    elif roll < 0.045:
        r["unit_price"] = 0
    elif roll < 0.06:
        r["unit_price"] = rng.choice(["N/A", "free", "TBD"])
    elif roll < 0.07:
        r["unit_price"] = 99999.99           # obvious typo / outlier
    elif roll < 0.15:
        r["unit_price"] = f"${price:,.2f}"   # e.g. "$1,299.00"

    # 9. Locations: wrong case and extra spaces
    if rng.random() < 0.20:
        r["city"] = messy_case(r["city"], rng)
    if rng.random() < 0.10:
        r["city"] = messy_spaces(r["city"], rng)
    if rng.random() < 0.15:
        r["state"] = r["state"].lower()
    if rng.random() < 0.02:
        r["city"] = ""

    return r


def make_duplicate(row: dict, rng: random.Random) -> dict:
    """
    Create a duplicate of an existing row. Half are exact copies, half are
    'near duplicates' that differ only in spacing/capitalisation - the kind a
    simple Excel 'Remove Duplicates' click would miss.
    """
    dup = dict(row)
    if rng.random() < 0.5:
        if dup["customer_name"]:
            dup["customer_name"] = messy_spaces(dup["customer_name"].upper(), rng)
        dup["city"] = messy_case(dup["city"], rng) if dup["city"] else dup["city"]
    return dup


def generate_rows(total_rows: int, seed: int, duplicate_share: float = 0.06) -> list[dict]:
    """Build the full list of messy rows (unique orders + duplicates), shuffled."""
    rng = random.Random(seed)  # fixed seed -> same file on every run
    customers = build_customers(rng)

    n_duplicates = int(total_rows * duplicate_share)
    n_unique = total_rows - n_duplicates

    rows = [
        add_quality_problems(make_clean_order(10001 + i, customers, rng), rng)
        for i in range(n_unique)
    ]
    # Pick rows to duplicate (only rows with an order ID, as real exports would)
    candidates = [r for r in rows if r["order_id"].strip()]
    rows += [make_duplicate(rng.choice(candidates), rng) for _ in range(n_duplicates)]

    rng.shuffle(rows)
    return rows


def save_csv(rows: list[dict], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    # newline="" prevents blank lines between rows on Windows
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a messy sales dataset.")
    parser.add_argument("--rows", type=int, default=500, help="Total rows (default 500)")
    parser.add_argument("--seed", type=int, default=42, help="Random seed (default 42)")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT,
                        help="Output CSV path (default data/raw_sales_data.csv)")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.rows < 20:
        print("ERROR: --rows must be at least 20.", file=sys.stderr)
        return 1
    try:
        rows = generate_rows(args.rows, args.seed)
        save_csv(rows, args.output)
    except PermissionError:
        print(f"ERROR: Cannot write to {args.output}. Is the file open in Excel? "
              "Close it and run again.", file=sys.stderr)
        return 1
    except OSError as exc:
        print(f"ERROR: Could not save the dataset: {exc}", file=sys.stderr)
        return 1

    print(f"Generated {len(rows)} messy sales records -> {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
