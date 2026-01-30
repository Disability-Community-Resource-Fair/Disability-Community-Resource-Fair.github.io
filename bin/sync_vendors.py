#!/usr/bin/env python3
import argparse
import csv
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
import unicodedata


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "_data"
POSTS_DIR = ROOT / "_posts"
TEMPLATE_MD = DATA_DIR / "2025-01-01-OrganizationName.md"


def strip_protocol(url: str) -> str:
    if not url:
        return url
    url = url.strip()
    url = re.sub(r"^https?://", "", url, flags=re.IGNORECASE)
    # collapse multiple slashes
    url = re.sub(r"/+", "/", url)
    # Keep trailing slash if in source, otherwise remove extra whitespace
    return url


def ascii_sanitize(text: str) -> str:
    if text is None:
        return ""
    text = str(text)
    # Normalize unicode accents to ASCII
    # text = unicodedata.normalize("NFKD", text)
    # text = text.encode("ascii", "ignore").decode("ascii")
    return text


def capitalize_sentences(text: str) -> str:
    # Capitalize the first alphabetic character after start/newline or sentence end punctuation.
    if not text:
        return text
    chars = list(text)
    start_sentence = True
    for i, c in enumerate(chars):
        if start_sentence and c.isalpha():
            chars[i] = c.upper()
            start_sentence = False
            continue
        if c in ".!?":
            start_sentence = True
        elif c == "\n":
            # Treat new lines as potential new sentences (for lists/paragraphs)
            start_sentence = True
    return "".join(chars)


def name_to_slug(name: str) -> str:
    name = ascii_sanitize(name)
    # Split on non-alphanumeric, title-case and join without separators to match sample (e.g., "1847Financial")
    parts = re.split(r"[^0-9A-Za-z]+", name)
    parts = [p for p in parts if p]
    if not parts:
        return "Vendor"
    # Preserve numbers as-is; title-case alpha tokens
    tokens = []
    for p in parts:
        if p.isdigit():
            tokens.append(p)
        else:
            tokens.append(p[:1].upper() + p[1:])
    return "".join(tokens)


def join_address(row: dict) -> str:
    a1 = (row.get("Street Address 1") or "").strip().rstrip(",")
    a2 = (row.get("Street Address 2") or "").strip()
    city = (row.get("City") or "").strip()
    state = (row.get("State") or "").strip()
    zipc = (row.get("Zip Code") or "").strip()
    line = a1
    if a2:
        if line:
            line += ", " + a2
        else:
            line = a2
    tail = ""
    if city:
        tail = city
    if state:
        tail = (tail + ", " if tail else "") + state
    if zipc:
        tail = (tail + " " if tail else "") + zipc
    if tail:
        line = (line + ", " if line else "") + tail
    return line


def normalize_items_to_underscored(items: list[str]) -> list[str]:
    out = []
    for it in items:
        t = ascii_sanitize((it or "").strip())
        if not t:
            continue
        # collapse internal whitespace to single spaces then replace with underscores
        t = re.sub(r"\s+", " ", t)
        t = t.replace(" ", "_")
        out.append(t)
    return out


def parse_csv(csv_path: Path) -> list[dict]:
    with csv_path.open(newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        rows = [dict(r) for r in reader]
    return rows

service_categories = {
"Behavioral Services, ABA Therapy, Assessment & Treatment",
"Social, Recreational, Fitness",
"Schools, Educational Programs, Advocacy",
"Transitioning Youth and Adults",
"Accessibility, Inclusion, Safety, Health",
"Camps",
"Therapeutic Services",
"Financial Planning, Insurance",
"Job Resources",
"Faith-based / Religious Organization",
}

@dataclass
class Vendor:
    name: str
    email: str
    phone: str
    website: str
    address: str
    description: str
    category: str        # single category (underscored)
    ages: list[str]      # tags

    @staticmethod
    def from_row(row: dict) -> "Vendor":
        name = (row.get("Organization Name") or "").strip()
        website = strip_protocol(row.get("Organization Website for handout") or "")
        email = (row.get("Organization Email for handout") or row.get("Contact Email Address") or "").strip()
        # Prefer organization phone for handout; fallback to contact phone
        phone = (row.get("Organization Phone Number for handout") or row.get("Contact Phone Number") or "").strip()
        address = join_address(row)
        description = (row.get("Brief description of organization for our handout: ") or "").replace("\r\n", "\n").replace("\r", "\n").strip()
        description = capitalize_sentences(description)

        raw_services = (row.get("Check the main type of service your organization provides. You may select more than one box if your organization provides a wide variety of services, but please only choose the categories that best describe the majority of the services you provide.") or "")
        # Treat the full string as a single category value, preserving commas but removing spaces by underscoring.
        categories = []
        for cat in service_categories:
            if cat in raw_services:
                categories.append(cat.replace(" ", "_"))
        category = " ".join(categories)

        raw_ages = (row.get("Check the age/grade range your organization serves. You may select more than one box if your organization provides a wide variety of services, but please only choose the categories that best describe the majority of the services you provide.") or "")
        ages_list = [s.strip() for s in raw_ages.split(",")]
        ages = normalize_items_to_underscored(ages_list)

        return Vendor(
            name=name,
            email=email,
            phone=phone,
            website=website,
            address=address,
            description=description,
            category=category,
            ages=ages,
        )


def load_template() -> str:
    if TEMPLATE_MD.exists():
        return TEMPLATE_MD.read_text(encoding='utf-8')
    # Fallback minimal template
    return (
        "---\n"
        "layout: post\n"
        "title: <OrganizationName>\n"
        "website: <OrganizationWebsite>\n"
        "tags: <OrganizationAgeRange>\n"
        "categories: <OrganizationService>\n"
        "address: <OrganizationAddress>\n"
        "phone_number: <OrganizationPhone>\n"
        "email: <OrganizationEmail>\n"
        "---\n"
        "<OrganizationDescription>\n"
    )


def date_prefix_from_template() -> str:
    # Infer date prefix from template filename: YYYY-MM-DD-OrganizationName.md
    m = re.match(r"(\d{4}-\d{2}-\d{2})-", TEMPLATE_MD.name)
    return m.group(1) if m else "2025-01-01"


def render_post(template: str, v: Vendor) -> str:
    # tags should be a space-separated list of tokens with underscores
    tags_val = " ".join(v.ages)
    # categories: single underscored token
    categories_val = v.category
    content = (
        template
        .replace("<OrganizationName>", v.name)
        .replace("<OrganizationWebsite>", v.website)
        .replace("<OrganizationAgeRange>", tags_val)
        .replace("<OrganizationService>", categories_val)
        .replace("<OrganizationAddress>", v.address)
        .replace("<OrganizationPhone>", v.phone)
        .replace("<OrganizationEmail>", v.email)
        .replace("<OrganizationDescription>", v.description)
    )
    return content


def extract_front_matter(md: str) -> tuple[dict, str]:
    # Simple front matter parser for key: value pairs (single-line values)
    if not md.startswith("---\n"):
        return {}, md
    parts = md.split("\n---\n", 1)
    if len(parts) < 2:
        return {}, md
    header = parts[0][4:]  # remove starting '---\n'
    body = parts[1]
    fm = {}
    for line in header.splitlines():
        if not line.strip():
            continue
        if ":" in line:
            k, v = line.split(":", 1)
            fm[k.strip()] = v.strip()
    return fm, body


def build_expected_fm_and_body(template: str, v: Vendor) -> tuple[dict, str]:
    rendered = render_post(template, v)
    fm, body = extract_front_matter(rendered)
    return fm, body


def load_if_exists(path: Path) -> str | None:
    if path.exists():
        return path.read_text(encoding='utf-8')
    return None


def update_existing_content(existing: str, expected_fm: dict, expected_body: str) -> str:
    fm, body = extract_front_matter(existing)
    # Update/insert keys
    for k, v in expected_fm.items():
        fm[k] = v
    new_header_lines = ["---"] + [f"{k}: {v}" for k, v in fm.items()] + ["---"]
    return "\n".join(new_header_lines) + "\n" + expected_body


def compute_filename(date_prefix: str, vendor_name: str) -> str:
    return f"{date_prefix}-{name_to_slug(vendor_name)}.md"


def sync(dry_run: bool = True, verbose: bool = False) -> int:
    csv_path = DATA_DIR / "vendors.csv"
    if not csv_path.exists():
        print(f"ERROR: CSV not found at {csv_path}", file=sys.stderr)
        return 1
    template = load_template()
    date_prefix = date_prefix_from_template()
    rows = parse_csv(csv_path)

    created, updated, skipped = 0, 0, 0
    MIN_DESC, MAX_DESC = 100, 800
    for row in rows:
        v = Vendor.from_row(row)
        if not v.name:
            if verbose:
                print("Skipping row with empty Organization Name")
            skipped += 1
            continue
        # Alert on description length outliers
        desc_len = len(v.description or "")
        if desc_len and desc_len < MIN_DESC:
            print(f"WARN: Short description for '{v.name}' ({desc_len} chars) < {MIN_DESC}")
        elif desc_len > MAX_DESC:
            print(f"WARN: Long description for '{v.name}' ({desc_len} chars) > {MAX_DESC}")
        fname = compute_filename(date_prefix, v.name)
        path = POSTS_DIR / fname
        expected_fm, expected_body = build_expected_fm_and_body(template, v)
        expected_content = "---\n" + "\n".join([f"{k}: {v}" for k, v in expected_fm.items()]) + "\n---\n" + expected_body

        existing = load_if_exists(path)
        if existing is None:
            created += 1
            if verbose:
                print(f"[NEW] {path.relative_to(ROOT)}")
            if not dry_run:
                path.write_text(expected_content, encoding='utf-8')
            continue

        # Compare and update if changed
        if existing != expected_content:
            updated += 1
            if verbose:
                print(f"[UPDATE] {path.relative_to(ROOT)}")
            if not dry_run:
                new_content = update_existing_content(existing, expected_fm, expected_body)
                path.write_text(new_content, encoding='utf-8')
        else:
            skipped += 1
            if verbose:
                print(f"[OK] {path.relative_to(ROOT)} (no changes)")

    print(f"Summary: created={created}, updated={updated}, unchanged={skipped}")
    return 0


def main():
    parser = argparse.ArgumentParser(description="Sync vendors CSV to _posts markdown files.")
    parser.add_argument("--write", action="store_true", help="Apply changes (default is dry-run)")
    parser.add_argument("--verbose", action="store_true", help="Verbose output")
    args = parser.parse_args()
    dry_run = not args.write
    return sync(dry_run=dry_run, verbose=args.verbose)


if __name__ == "__main__":
    raise SystemExit(main())
