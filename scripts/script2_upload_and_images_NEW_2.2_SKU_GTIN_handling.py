import os
import json
import time
import requests
import pandas as pd
import base64
import sys
import msvcrt
from openai import OpenAI
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from rapidfuzz import process, fuzz

#########################################
# PATHS
#########################################

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_PATH = os.path.dirname(SCRIPT_DIR)

print(f"Using BASE_PATH: {BASE_PATH}")

INPUT_FILE = os.path.join(BASE_PATH, "working", "products_enriched.xlsx")

IMAGE_LIBRARY = os.path.join(BASE_PATH, "image_library")
GENERATED_IMAGES = os.path.join(IMAGE_LIBRARY, "generated")
LOGS_DIR = os.path.join(BASE_PATH, "logs")

# IMPORTANT: keep original structure
CSV_LIBRARY = os.path.join(IMAGE_LIBRARY, "image_library_V2.0.csv")

SKIPPED_LOG = os.path.join(LOGS_DIR, "skipped_images.csv")
MISSING_IMAGE_REVIEW_LOG = os.path.join(LOGS_DIR, "missing_image_review.csv")
#IMAGES FOUND BUT NOT USED NEED TO CHECK LATER
FOUND_NOT_USED_LOG = os.path.join(LOGS_DIR, "found_not_used_images.csv")

#SAVE STATE THE SCRIPT CAN BE RE-RUN WITHOUT DUPLICATING WORK
RUN_STATE_FILE = os.path.join(LOGS_DIR, "script2_run_state.json")

NIELSEN_NORMALIZED_FILE = os.path.join(BASE_PATH, "data", "nielsen_normalized.csv")

TEMPLATE_IMAGE_VALIDATION = os.path.join(BASE_PATH, "templates", "image_validation_prompt.txt")
TEMPLATE_SINGLE_IMAGE_VALIDATION = os.path.join(BASE_PATH, "templates", "generated_image_validation_prompt.txt")


#########################################
# CREATE FOLDERS
#########################################

os.makedirs(IMAGE_LIBRARY, exist_ok=True)
os.makedirs(GENERATED_IMAGES, exist_ok=True)
os.makedirs(LOGS_DIR, exist_ok=True)

#########################################
# OPENAI CLIENT
#########################################

client = OpenAI()


#########################################
# RUN OPTIONS
#########################################

def ask_ai_image_toggle():
    while True:
        choice = input("Do you want to use AI image generation for this run? (yes/no): ").strip().lower()

        if choice in ["yes", "y"]:
            print("AI image generation: ENABLED")
            return True
        elif choice in ["no", "n"]:
            print("AI image generation: DISABLED")
            return False
        else:
            print("Please type yes or no.")

#########################################
#NEW STARTUP PROMPT TO DECIDE WHETHER TO USE AI IMAGE GENERATION OR NOT. THIS WAY WE CAN CONTROL THIS FEATURE WITHOUT HAVING TO CHANGE CODE OR RE-EXPORT EXCEL FILES WITH DIFFERENT FLAGS.
#########################################
def ask_use_image_library():
    while True:
        choice = input("Do you want to use the image library for this run? (yes/no): ").strip().lower()

        if choice in ["yes", "y"]:
            print("Image library usage: ENABLED")
            return True
        elif choice in ["no", "n"]:
            print("Image library usage: DISABLED")
            return False
        else:
            print("Please type yes or no.")


def ask_use_generated_library_images():
    while True:
        choice = input("Do you want to use generated images from the image library? (yes/no): ").strip().lower()

        if choice in ["yes", "y"]:
            print("Generated library image usage: ENABLED")
            return True
        elif choice in ["no", "n"]:
            print("Generated library image usage: DISABLED")
            return False
        else:
            print("Please type yes or no.")
#########################################
# CHECK DIGIT PROMPT
#########################################
def ask_check_digit_setting():
    while True:
        choice = input("Does your POS keep check digit? (yes/no): ").strip().lower()

        if choice in ["yes", "y"]:
            print("Check digit handling: KEEP ORIGINAL")
            return True
        elif choice in ["no", "n"]:
            print("Check digit handling: RE-CALCULATE / NORMALIZE")
            return False
        else:
            print("Please type yes or no.")

#########################################
# RUN STATE / RESUME / SAFE STOP
#########################################

def load_run_state():
    if not os.path.exists(RUN_STATE_FILE):
        return None

    try:
        with open(RUN_STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def save_run_state(
    last_completed_index,
    total_products,
    last_completed_sku="",
    last_completed_product_name="",
    status="running"
):
    state = {
        "company_name": company_name,
        "input_file": INPUT_FILE,
        "total_products": total_products,
        "last_completed_index": last_completed_index,
        "last_completed_sku": str(last_completed_sku),
        "last_completed_product_name": str(last_completed_product_name),
        "status": status,
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")
    }

    with open(RUN_STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def reset_run_state():
    if os.path.exists(RUN_STATE_FILE):
        os.remove(RUN_STATE_FILE)


def ask_resume_if_available(total_products):
    state = load_run_state()

    if not state:
        return 0

    if state.get("status") == "complete":
        return 0

    saved_company = state.get("company_name", "")
    saved_input = state.get("input_file", "")
    saved_index = int(state.get("last_completed_index", -1))
    saved_total = int(state.get("total_products", 0))

    if saved_company != company_name or saved_input != INPUT_FILE:
        print("Existing run state belongs to a different company or input file.")
        print("Ignoring old state and starting from the beginning.")
        return 0

    if saved_index < 0:
        return 0

    resume_from = saved_index + 1

    if resume_from >= total_products:
        return 0

    while True:
        choice = input(
            f"Resume from {resume_from}/{total_products}? (yes/no): "
        ).strip().lower()

        if choice in ["yes", "y"]:
            print(f"Resuming from row {resume_from}/{total_products}")
            return resume_from

        elif choice in ["no", "n"]:
            reset_run_state()
            print("Starting from the beginning. Previous save state cleared.")
            return 0

        else:
            print("Please type yes or no.")


def esc_pressed():
    try:
        if msvcrt.kbhit():
            key = msvcrt.getch()
            if key == b'\x1b':
                return True
    except Exception:
        pass
    return False


def stop_requested():
    if esc_pressed():
        print("\n[STOP REQUESTED] Escape pressed. Script will stop safely after this checkpoint.")
        return True
    return False


#########################################
# LOAD CONFIG
#########################################

CONFIG_PATH = os.path.join(BASE_PATH, "config", "company_config.json")

with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    config = json.load(f)

company_name = config["company_name"]
company_url = config["company_url"]

BASE_URL = company_url
PRODUCTS_URL = f"{BASE_URL}/admin/products"


#########################################
# IMAGE LIBRARY LOOKUP
#########################################

def load_image_library():
    if not os.path.exists(CSV_LIBRARY):
        return pd.DataFrame()

    try:
        df = pd.read_csv(CSV_LIBRARY, dtype=str)
        df.columns = [str(col).strip() for col in df.columns]
        df = df.fillna("")
        return df
    except Exception as e:
        print(f"Could not load image library CSV: {e}")
        return pd.DataFrame()


def clean_match_text(text):
    return clean_text(text)


def image_source_is_generated(source_value):
    return str(source_value).strip().lower() == "openai_generated"


def find_image_in_library(normalized_name, product_name, image_keyword, use_generated_library_images):
    if image_library_df.empty:
        return None

    df = image_library_df.copy()

    required_cols = [
        "product_name",
        "sku",
        "company_name",
        "image_source",
        "location",
        "normalized_name",
        "image_keyword"
    ]

    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    if not use_generated_library_images:
        df = df[
            ~df["image_source"].astype(str).str.strip().str.lower().eq("openai_generated")
        ].copy()

    if df.empty:
        return None

    df["product_name_clean"] = df["product_name"].astype(str).apply(clean_match_text)
    df["normalized_name_clean"] = df["normalized_name"].astype(str).apply(clean_match_text)
    df["image_keyword_clean"] = df["image_keyword"].astype(str).apply(clean_match_text)

    normalized_name_clean = clean_match_text(normalized_name)
    product_name_clean = clean_match_text(product_name)
    image_keyword_clean = clean_match_text(image_keyword)

    def build_result(row, matched_on):
        location = str(row.get("location", "")).strip()

        if location and os.path.exists(location):
            return {
                "source": f"image_library_{row.get('image_source', 'unknown')}",
                "local_path": location,
                "image_url": "",
                "matched_on": matched_on
            }

        return None

    # Priority 1 — normalized_name
    if normalized_name_clean:
        match_df = df[df["normalized_name_clean"] == normalized_name_clean]
        if not match_df.empty:
            result = build_result(match_df.iloc[0], "normalized_name")
            if result:
                return result

    # Priority 2 — product_name
    if product_name_clean:
        match_df = df[df["product_name_clean"] == product_name_clean]
        if not match_df.empty:
            result = build_result(match_df.iloc[0], "product_name")
            if result:
                return result

    # Priority 3 — image_keyword
    if image_keyword_clean:
        match_df = df[df["image_keyword_clean"] == image_keyword_clean]
        if not match_df.empty:
            result = build_result(match_df.iloc[0], "image_keyword")
            if result:
                return result

    return None



#########################################
# LOAD DATA
#########################################

products = pd.read_excel(INPUT_FILE, dtype={"sku": str})
products["sku"] = products["sku"].str.zfill(13)

if os.path.exists(NIELSEN_NORMALIZED_FILE):
    nielsen_normalized = pd.read_csv(NIELSEN_NORMALIZED_FILE)
    nielsen_normalized.columns = nielsen_normalized.columns.str.strip()

else:
    nielsen_normalized = pd.DataFrame()

nielsen_count = 0
pexels_count = 0
generated_count = 0
#########################################
# SELENIUM SETUP
#########################################

chrome_options = Options()
chrome_options.add_argument("--start-maximized")

service = Service()

driver = webdriver.Chrome(service=service, options=chrome_options)

wait = WebDriverWait(driver, 20)


#########################################
# LOGIN STEP
#########################################

driver.get(BASE_URL + "/admin")

print("Please log in to GrazeCart admin in the opened browser.")
input("Press ENTER here after login is complete...")

#########################################
# PEXELS API
#########################################

PEXELS_API_KEY = "YOUR_PEXELS_API_KEY"

def get_pexels_image(keyword):

    url = "https://api.pexels.com/v1/search"

    headers = {
        "Authorization": PEXELS_API_KEY
    }

    params = {
        "query": keyword,
        "per_page": 1
    }

    r = requests.get(url, headers=headers, params=params)

    if r.status_code != 200:
        return None

    data = r.json()

    if not data["photos"]:
        return None

    return data["photos"][0]["src"]["large"]

#########################################
# NIELSEN IMAGE
#########################################


#########################################
# NIELSEN IMAGE MATCHING (FAST + STRONG)
#########################################

import re
from collections import defaultdict, Counter
from rapidfuzz import fuzz

STOPWORDS = {
    "the", "a", "an", "and", "or", "with", "for",
    "pack", "package", "pkg", "item", "product",
    "food", "grocery", "fresh"
}

LOW_WEIGHT_TOKENS = {
    "oz", "fl", "lb", "lbs", "g", "kg", "ml", "l",
    "ct", "count", "pk", "qt", "pt", "gal"
}

def clean_text(text):
    text = str(text).lower().strip()

    text = text.replace("&", " and ")
    text = text.replace("/", " ")
    text = text.replace("-", " ")

    text = re.sub(r"(\d+)(oz|lb|lbs|g|kg|ml|l|ct|pk)\b", r"\1 \2", text)
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    return text

#########################################
# UPC / GTIN HELPERS
#########################################

def digits_only(value):
    return re.sub(r"\D", "", str(value))


def calculate_upc_check_digit(upc_without_check_digit):
    code = digits_only(upc_without_check_digit)

    if len(code) != 11:
        return None

    total = 0
    for i, ch in enumerate(code):
        digit = int(ch)
        total += digit * 3 if i % 2 == 0 else digit

    return str((10 - (total % 10)) % 10)


def calculate_ean13_check_digit(ean_without_check_digit):
    code = digits_only(ean_without_check_digit)

    if len(code) != 12:
        return None

    total = 0
    for i, ch in enumerate(code):
        digit = int(ch)
        total += digit if i % 2 == 0 else digit * 3

    return str((10 - (total % 10)) % 10)


def expand_upce_to_upca(upce):
    code = digits_only(upce)

    if len(code) == 6:
        number_system = "0"
        payload = code
        supplied_check = None
    elif len(code) == 7:
        number_system = code[0]
        payload = code[1:]
        supplied_check = None
    elif len(code) == 8:
        number_system = code[0]
        payload = code[1:7]
        supplied_check = code[7]
    else:
        return None

    if number_system not in ["0", "1"]:
        return None

    d1, d2, d3, d4, d5, d6 = payload

    if d6 in ["0", "1", "2"]:
        upc_body_11 = number_system + d1 + d2 + d6 + "0000" + d3 + d4 + d5
    elif d6 == "3":
        upc_body_11 = number_system + d1 + d2 + d3 + "00000" + d4 + d5
    elif d6 == "4":
        upc_body_11 = number_system + d1 + d2 + d3 + d4 + "00000" + d5
    else:
        upc_body_11 = number_system + d1 + d2 + d3 + d4 + d5 + "0000" + d6

    check_digit = calculate_upc_check_digit(upc_body_11)

    if check_digit is None:
        return None

    upca = upc_body_11 + check_digit

    if supplied_check and supplied_check != check_digit:
        return None

    return upca


def generate_barcode_candidates(raw_sku, keep_check_digit=True):
    raw = digits_only(raw_sku)
    candidates = []

    def add(value, label):
        value = digits_only(value)
        if value and value not in [c["value"] for c in candidates]:
            candidates.append({"value": value, "label": label})

    if not raw:
        return candidates

    add(raw, "raw_sku_digits")

    if len(raw) == 13:
        if keep_check_digit:
            add(raw, "gtin13_keep_check_digit")
            if raw.startswith("0"):
                add(raw[1:], "upca_from_gtin13")
        else:
            if raw.startswith("00"):
                base_11 = raw[-11:]
                check = calculate_upc_check_digit(base_11)
                if check is not None:
                    add(base_11 + check, "upca_recalculated_from_13_no_check")
                    add("0" + base_11 + check, "gtin13_from_upca_recalculated_no_check")
            elif raw.startswith("0"):
                base_12 = raw[1:13]
                check = calculate_ean13_check_digit(base_12)
                if check is not None:
                    add(base_12 + check, "ean13_recalculated_no_check")

    elif len(raw) == 12:
        if keep_check_digit:
            add(raw, "upca_keep_check_digit")
            add("0" + raw, "gtin13_from_upca")
        else:
            base_11 = raw[:11]
            check = calculate_upc_check_digit(base_11)
            if check is not None:
                add(base_11 + check, "upca_recalculated_no_check")
                add("0" + base_11 + check, "gtin13_from_upca_recalculated_no_check")

    elif len(raw) in [6, 7, 8]:
        upca = expand_upce_to_upca(raw)

        if upca:
            add(upca, "upca_from_upce")
            add("0" + upca, "gtin13_from_upce")

        if len(raw) == 8:
            add(raw, "ean8_direct")

    return candidates


def tokenize(text):
    return clean_text(text).split()

def useful_tokens(tokens):
    return [t for t in tokens if t not in STOPWORDS]

def token_weight(token, token_freq):
    if token in LOW_WEIGHT_TOKENS:
        return 0.35
    freq = token_freq.get(token, 1)
    if freq <= 2:
        return 2.5
    if freq <= 5:
        return 2.0
    if freq <= 20:
        return 1.5
    if freq <= 100:
        return 1.0
    return 0.6

def build_nielsen_index(df):
    df = df.copy()

    df.columns = df.columns.str.strip()

    df = df.dropna(subset=["canonical_name", "Shot1Link"]).copy()
    df = df[
        (df["canonical_name"].astype(str).str.strip() != "") &
        (df["Shot1Link"].astype(str).str.strip() != "")
    ].copy()

    df["canonical_clean"] = df["canonical_name"].apply(clean_text)
    df["canonical_tokens"] = df["canonical_clean"].apply(tokenize)
    df["canonical_token_set"] = df["canonical_tokens"].apply(set)

    token_freq = Counter()
    for tokens in df["canonical_token_set"]:
        for token in tokens:
            if token not in STOPWORDS:
                token_freq[token] += 1

    inverted_index = defaultdict(set)
    for idx, tokens in df["canonical_token_set"].items():
        for token in tokens:
            if token not in STOPWORDS:
                inverted_index[token].add(idx)

    return df, inverted_index, token_freq

# BUILD INDEX (runs once)
if not nielsen_normalized.empty:
    nielsen_normalized, nielsen_inverted_index, nielsen_token_freq = build_nielsen_index(nielsen_normalized)
else:
    nielsen_inverted_index = defaultdict(set)
    nielsen_token_freq = Counter()


#########################################
# NIELSEN UPC / GTIN INDEX
#########################################

def build_nielsen_barcode_index(df):
    barcode_col = "GTIN"

    if barcode_col not in df.columns:
        print("GTIN column not found in Nielsen file. Barcode search disabled.")
        return {}, None

    index = {}

    for idx, value in df[barcode_col].items():
        code = digits_only(value)

        if not code:
            continue

        possible_values = set()
        possible_values.add(code)

        if len(code) == 12:
            possible_values.add("0" + code)
        elif len(code) == 13 and code.startswith("0"):
            possible_values.add(code[1:])

        for candidate in possible_values:
            if candidate not in index:
                index[candidate] = idx

    print(f"Nielsen barcode search enabled using column: {barcode_col}")
    return index, barcode_col


if not nielsen_normalized.empty:
    nielsen_barcode_index, nielsen_barcode_column = build_nielsen_barcode_index(nielsen_normalized)
else:
    nielsen_barcode_index = {}
    nielsen_barcode_column = None


def get_nielsen_match_by_barcode(raw_sku, keep_check_digit=True):
    if not nielsen_barcode_index:
        return None

    candidates = generate_barcode_candidates(raw_sku, keep_check_digit=keep_check_digit)

    for candidate in candidates:
        value = candidate["value"]

        if value in nielsen_barcode_index:
            idx = nielsen_barcode_index[value]
            row = nielsen_normalized.loc[idx]

            return {
                "image_url": row["Shot1Link"],
                "canonical_name": row["canonical_name"],
                "canonical_clean": row.get("canonical_clean", clean_text(row["canonical_name"])),
                "score": 1.0,
                "exact_match": True,
                "match_type": "barcode",
                "barcode_candidate": value,
                "barcode_candidate_label": candidate["label"],
                "barcode_column": nielsen_barcode_column
            }

    return None


def get_candidate_indices(search_text, max_candidates=400):
    tokens = useful_tokens(tokenize(search_text))

    if not tokens:
        return []

    ranked_tokens = sorted(
        set(tokens),
        key=lambda t: (nielsen_token_freq.get(t, 999999), -len(t))
    )

    candidate_counter = Counter()

    for token in ranked_tokens:
        matches = nielsen_inverted_index.get(token, set())

        for idx in matches:
            candidate_counter[idx] += token_weight(token, nielsen_token_freq)

    return [idx for idx, _ in candidate_counter.most_common(max_candidates)]

def score_candidate(search_text, row):
    search_clean = clean_text(search_text)
    search_tokens = useful_tokens(tokenize(search_text))
    search_set = set(search_tokens)

    candidate_clean = row["canonical_clean"]
    candidate_set = row["canonical_token_set"]

    if not search_set or not candidate_set:
        return 0

    shared = search_set & candidate_set
    if not shared:
        return 0

    shared_weight = sum(token_weight(t, nielsen_token_freq) for t in shared)
    input_weight = sum(token_weight(t, nielsen_token_freq) for t in search_set) or 1
    candidate_weight = sum(token_weight(t, nielsen_token_freq) for t in candidate_set) or 1

    weighted_input_coverage = shared_weight / input_weight
    weighted_candidate_coverage = shared_weight / candidate_weight

    token_set_score = fuzz.token_set_ratio(search_clean, candidate_clean) / 100.0
    token_sort_score = fuzz.token_sort_ratio(search_clean, candidate_clean) / 100.0

    final_score = (
        0.4 * weighted_input_coverage +
        0.2 * weighted_candidate_coverage +
        0.25 * token_set_score +
        0.15 * token_sort_score
    )

    return final_score

def get_nielsen_match(search_name, min_score=0.62):

    if nielsen_normalized.empty:
        return None

    search_clean = clean_text(search_name)
    if not search_clean:
        return None

    candidates = get_candidate_indices(search_name)

    best_score = 0
    best_idx = None

    for idx in candidates:
        row = nielsen_normalized.loc[idx]
        score = score_candidate(search_name, row)

        if score > best_score:
            best_score = score
            best_idx = idx

    if best_idx is None or best_score < min_score:
        return None

    best_row = nielsen_normalized.loc[best_idx]

    return {
        "image_url": best_row["Shot1Link"],
        "canonical_name": best_row["canonical_name"],
        "canonical_clean": best_row["canonical_clean"],
        "score": best_score,
        "exact_match": search_clean == best_row["canonical_clean"]
    }


#########################################
# IMAGE VALIDATION
#########################################

def validate_images(product_name, img1, img2):

    with open(TEMPLATE_IMAGE_VALIDATION, "r", encoding="utf-8") as f:
        template = f.read()

    prompt = template.format(product_name=product_name)

    content = [
        {
            "type": "input_text",
            "text": prompt
        }
    ]

    if img1:
        content.append({
            "type": "input_image",
            "image_url": img1
        })

    if img2:
        content.append({
            "type": "input_image",
            "image_url": img2
        })

    response = client.responses.create(
        model="gpt-5",
        input=[
            {
                "role": "user",
                "content": content
            }
        ]
    )

    text = response.output_text

    lines = text.split("\n")

    try:

        image1_score = int(lines[0].split(":")[1].strip())
        image2_score = int(lines[1].split(":")[1].strip())
        best_image = lines[2].split(":")[1].strip()
        decision = lines[3].split(":")[1].strip()

        return image1_score, image2_score, best_image, decision

    except:
        return 0, 0, "NONE", "GENERATE"

#########################################
# IMAGE VALIDATION NEW
#########################################

def validate_single_image(product_name, img1):

    with open(TEMPLATE_SINGLE_IMAGE_VALIDATION, "r", encoding="utf-8") as f:
        template = f.read()

    prompt = template.format(product_name=product_name)

    content = [
        {
            "type": "input_text",
            "text": prompt
        },
        {
            "type": "input_image",
            "image_url": img1
        }
    ]

    response = client.responses.create(
        model="gpt-5-mini",
        input=[
            {
                "role": "user",
                "content": content
            }
        ]
    )

    text = response.output_text
    lines = text.split("\n")

    try:
        score = int(lines[0].split(":")[1].strip())
        decision = lines[1].split(":")[1].strip()
        return score, decision
    except:
        return 0, "REGENERATE"



#########################################
# SAFE FILE NAMES
#########################################

def make_safe_filename(text):
    text = str(text).lower().strip()

    # replace spaces with underscores
    text = text.replace(" ", "_")

    # replace Windows-invalid filename characters
    for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
        text = text.replace(ch, "_")

    # collapse repeated underscores
    while "__" in text:
        text = text.replace("__", "_")

    return text.strip("_")


#########################################
# IMAGE GENERATION
#########################################


def generate_image(product_name, keyword):

    prompt = f"""
professional grocery product photograph of {product_name}
clean white background
studio lighting
realistic product photography
"""

    img = client.images.generate(
        model="gpt-image-1",
        prompt=prompt,
        size="1024x1024"
    )

    data = img.data[0]

    url = getattr(data, "url", None)
    b64 = getattr(data, "b64_json", None)

    clean_name = make_safe_filename(product_name)
    clean_keyword = make_safe_filename(keyword)

    filename = f"{clean_name}_{clean_keyword}_openai.png"
    filepath = os.path.join(GENERATED_IMAGES, filename)

    #########################################
    # CASE 1 — URL RETURNED
    #########################################

    if url:
        print("Image returned as URL")

        r = requests.get(url)

        with open(filepath, "wb") as f:
            f.write(r.content)

        return filepath

    #########################################
    # CASE 2 — BASE64 RETURNED
    #########################################

    if b64:
        print("Image returned as BASE64")

        image_bytes = base64.b64decode(b64)

        with open(filepath, "wb") as f:
            f.write(image_bytes)

        return filepath

    #########################################
    # CASE 3 — FAILURE
    #########################################

    print("Image generation failed — no URL or base64 returned")

    return None


#########################################
# DOWNLOAD IMAGE (PEXELS)
#########################################

def download_image(url, normalized_name, product_name, company_name, source):

    clean_normalized_name = make_safe_filename(normalized_name)
    clean_product_name = make_safe_filename(product_name)
    clean_company_name = make_safe_filename(company_name)

    filename = f"{clean_normalized_name}_{clean_product_name}_{clean_company_name}_{source}.jpg"
    filepath = os.path.join(IMAGE_LIBRARY, filename)

    if os.path.exists(filepath):
        print(f"Image already exists in library, reusing: {filepath}")
        return filepath

    r = requests.get(url)

    with open(filepath, "wb") as f:
        f.write(r.content)

    return filepath

#########################################
# IMAGE LIBRARY TRACKING
#########################################

def record_image(product_name, normalized_name, image_keyword, sku, source, local_path, image_url=""):

    row = {
        "product_name": product_name,
        "sku": sku,
        "company_name": company_name,
        "image_source": source,
        "location": local_path,
        "normalized_name": normalized_name,
        "image_keyword": image_keyword
    }

    df = pd.DataFrame([row])

    if os.path.exists(CSV_LIBRARY):
        df.to_csv(CSV_LIBRARY, mode="a", header=False, index=False)
    else:
        df.to_csv(CSV_LIBRARY, index=False)

def record_skipped(product_name, sku, reason):

    row = {
        "product_name": product_name,
        "sku": sku,
        "reason": reason
    }

    df = pd.DataFrame([row])

    if os.path.exists(SKIPPED_LOG):
        df.to_csv(SKIPPED_LOG, mode="a", header=False, index=False)
    else:
        df.to_csv(SKIPPED_LOG, index=False)
#########################################
# MISSING IMAGE REVIEW LOG
#########################################

def record_missing_image_review(
    product_name,
    sku,
    keyword,
    normalized_name,
    reason,
    company_name,
    ai_generation_enabled,
    nielsen_match=None,
    nielsen_img=None,
    pexels_img=None,
    validation_decision="",
    best_image="",
    image1_score="",
    image2_score="",
    single_image_score="",
    single_image_decision=""
):
    row = {
        "product_name": product_name,
        "sku": sku,
        "image_keyword": keyword,
        "normalized_name": normalized_name,
        "company_name": company_name,
        "ai_generation_enabled": ai_generation_enabled,
        "reason": reason,

        "nielsen_found": bool(nielsen_match),
        "nielsen_score": nielsen_match.get("score") if nielsen_match else "",
        "nielsen_exact_match": nielsen_match.get("exact_match") if nielsen_match else "",
        "nielsen_canonical_name": nielsen_match.get("canonical_name") if nielsen_match else "",
        "nielsen_canonical_clean": nielsen_match.get("canonical_clean") if nielsen_match else "",
        "nielsen_image_url": nielsen_img if nielsen_img else "",

        "pexels_found": bool(pexels_img),
        "pexels_image_url": pexels_img if pexels_img else "",

        "validation_decision": validation_decision,
        "best_image": best_image,
        "image1_score": image1_score,
        "image2_score": image2_score,
        "single_image_score": single_image_score,
        "single_image_decision": single_image_decision
    }

    df = pd.DataFrame([row])

    if os.path.exists(MISSING_IMAGE_REVIEW_LOG):
        df.to_csv(MISSING_IMAGE_REVIEW_LOG, mode="a", header=False, index=False)
    else:
        df.to_csv(MISSING_IMAGE_REVIEW_LOG, index=False)


#########################################
# FOUND BUT NOT USED IMAGE LOG
#########################################

def record_found_not_used(
    product_name,
    sku,
    keyword,
    normalized_name,
    company_name,
    found_source,
    found_image_url,
    reason,
    validation_decision="",
    best_image="",
    image1_score="",
    image2_score="",
    single_image_score="",
    single_image_decision="",
    nielsen_score="",
    nielsen_exact_match=""
):
    row = {
        "product_name": product_name,
        "sku": sku,
        "image_keyword": keyword,
        "normalized_name": normalized_name,
        "company_name": company_name,
        "found_source": found_source,
        "found_image_url": found_image_url,
        "reason_not_used": reason,
        "validation_decision": validation_decision,
        "best_image": best_image,
        "image1_score": image1_score,
        "image2_score": image2_score,
        "single_image_score": single_image_score,
        "single_image_decision": single_image_decision,
        "nielsen_score": nielsen_score,
        "nielsen_exact_match": nielsen_exact_match
    }

    df = pd.DataFrame([row])

    if os.path.exists(FOUND_NOT_USED_LOG):
        df.to_csv(FOUND_NOT_USED_LOG, mode="a", header=False, index=False)
    else:
        df.to_csv(FOUND_NOT_USED_LOG, index=False)


#########################################
# SELENIUM SELECTORS
#########################################

def click_save(driver):

    save_btn = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//button[contains(text(),'Save')]"))
    )

    driver.execute_script("arguments[0].scrollIntoView(true);", save_btn)

    driver.execute_script("arguments[0].click();", save_btn)

    try:
        WebDriverWait(driver, 10).until_not(
            EC.presence_of_element_located((By.CLASS_NAME, "notification__content"))
        )
    except:
        time.sleep(2)


#########################################
# PRODUCT FIELD HANDLERS
#########################################

def process_input_field(name_attr, excel_value):

    if pd.isna(excel_value) or str(excel_value).strip().lower() in ["", "nan", "none", "null"]:
        return

    field = driver.find_element(By.NAME, name_attr)

    field.clear()

    field.send_keys(str(excel_value).strip())


def process_redactor(name_attr, excel_value):

    if pd.isna(excel_value) or str(excel_value).strip().lower() in ["", "nan", "none", "null"]:
        return

    textarea = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.NAME, name_attr))
    )

    value = str(excel_value).strip()

    driver.execute_script("""
        const el = arguments[0];
        const val = arguments[1];

        el.value = val;
        el.innerHTML = val;

        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur', { bubbles: true }));
    """, textarea, value)

    time.sleep(0.5)


def process_dropdown(name_attr, excel_value):

    if pd.isna(excel_value) or str(excel_value).strip().lower() in ["", "nan", "none", "null"]:
        return

    select = Select(driver.find_element(By.NAME, name_attr))

    excel_value = str(excel_value).strip().lower()

    if name_attr == "inventory_type":

        packing_map = {
            "frozen": "1",
            "dry": "2",
            "fresh": "3",
            "seasonal": "4",
            "digital": "6"
        }

        excel_value = packing_map.get(excel_value, excel_value)

    elif name_attr == "track_inventory":

        track_map = {
            "track inventory": "yes",
            "track inventory on bundle": "bundle",
            "do not track inventory": "no"
        }

        excel_value = track_map.get(excel_value, excel_value)

    select.select_by_value(excel_value)


def process_radio(name_attr, excel_value):

    if not excel_value:
        return

    value = "1" if str(excel_value).strip().lower() == "yes" else "0"

    xpath_options = [
        f"//input[@name='{name_attr}' and @value='{value}']",
        f"//input[contains(@name, '{name_attr}') and @value='{value}']"
    ]

    last_error = None

    for xpath in xpath_options:
        try:
            radio = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", radio)
            time.sleep(0.5)

            try:
                radio.click()
            except:
                driver.execute_script("arguments[0].click();", radio)

            return

        except Exception as e:
            last_error = e

    print(f"Could not find radio field: name={name_attr}, value={value}")
    raise last_error



#########################################
# SKU SEARCH
#########################################

def search_by_sku(driver, sku):

    print("Opening products page...")

    driver.get(PRODUCTS_URL)

    # Wait for search form to load
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, "filterproductsForm"))
    )

    search_box = driver.find_element(
        By.CSS_SELECTOR,
        "#filterproductsForm input[name='products']"
    )

    search_box.clear()
    search_box.send_keys(str(sku))

    print(f"Searching SKU: {sku}")

    search_button = driver.find_element(
        By.CSS_SELECTOR,
        "#filterproductsForm button[type='submit'].btn-input-overlay"
    )

    search_button.click()

    time.sleep(2)
    # Wait for results table
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//table//tbody//tr"))
        )
    except:
        print(f"SKU NOT FOUND: {sku}")
        return False

    rows = driver.find_elements(By.XPATH, "//table//tbody//tr")

    for row in rows:

        try:

            sku_cell = row.find_element(By.XPATH, ".//td[@data-label='SKU']")
            sku_text = sku_cell.text.strip()

            if sku_text == sku:

                product_link = row.find_element(By.XPATH, ".//td[1]//a")

                driver.execute_script(
                    "arguments[0].click();",
                    product_link
                )

                # Wait for product page to load fully
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located(
                        (By.NAME, "unit_description")
                    )
                )

                print("Product opened.")

                return True

        except:
            continue

    print(f"SKU NOT FOUND IN RESULTS: {sku}")

    return False



#########################################
# PRODUCT DATA UPLOAD
#########################################


#########################################
# TAB NAVIGATION HELPERS
#########################################

def open_product_tab(tab_name, retries=3, delay=2):
    base_url = driver.current_url.split("?")[0]
    target_url = f"{base_url}?tab={tab_name}"

    last_error = None

    for attempt in range(1, retries + 1):
        try:
            driver.get(target_url)

            WebDriverWait(driver, 20).until(
                lambda d: f"tab={tab_name}" in d.current_url
            )

            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "productForm"))
            )

            time.sleep(2)
            return

        except Exception as e:
            last_error = e
            print(f"Retry opening tab '{tab_name}' ({attempt}/{retries})")
            time.sleep(delay)

    raise last_error


#########################################
# PRODUCT DATA UPLOAD
#########################################

def upload_product_data(row):

    print("Uploading product:", row["product_name"])

    #########################################
    # DESCRIPTION TAB
    #########################################

    print("Updating description fields...")

    process_input_field("unit_description", row["unit_description"])
    process_redactor("summary", row["summary"])
    process_redactor("description", row["description"])
    process_input_field("ingredients", row["ingredients"])

    click_save(driver)

    #########################################
    # PRICE TAB
    #########################################

    open_product_tab("price")

    # weight and price section skipped as it's managed by ITR
    click_save(driver)

    #########################################
    # INVENTORY TAB
    #########################################

    open_product_tab("inventory")

    process_radio("is_bundle", row["bundle_product"])
    process_dropdown("track_inventory", row["track_inventory"])
    process_input_field("stock_out_inventory", row["stockout_threshold"])
    process_radio("back_order", row["allow_backorder"])

    click_save(driver)

    #########################################
    # SETTINGS TAB
    #########################################

    open_product_tab("settings")

    process_radio("visible", row["visibility"])
    process_radio("taxable", row["taxable"])
    process_dropdown("inventory_type", row["packing_group"])

    process_input_field(
        "settings[standard_callout_message]",
        row["standard_callout"]
    )

    process_input_field(
        "settings[sale_message]",
        row["sale_callout"]
    )

    click_save(driver)

    #########################################
    # SEO TAB
    #########################################

    open_product_tab("seo")

    process_input_field("page_title", row["meta_title"])
    process_input_field("seo_description", row["meta_description"])
    process_radio("seo_visibility", row["visible_to_search"])

    click_save(driver)


#########################################
# IMAGE UPLOAD
#########################################

def upload_image(image_path):

    print("Uploading image:", image_path)

    upload_input = driver.find_element(By.XPATH, "//input[@type='file']")

    upload_input.send_keys(image_path)

    time.sleep(6)

    click_save(driver)





#########################################
# MAIN LOOP
#########################################

USE_AI_IMAGE_GEN = ask_ai_image_toggle()
KEEP_CHECK_DIGIT = ask_check_digit_setting()
USE_IMAGE_LIBRARY = ask_use_image_library()
USE_GENERATED_LIBRARY_IMAGES = ask_use_generated_library_images() if USE_IMAGE_LIBRARY else False

image_library_df = load_image_library() if USE_IMAGE_LIBRARY else pd.DataFrame()

total_products = len(products)
start_index = ask_resume_if_available(total_products)

for i in range(start_index, total_products):

    if stop_requested():
        save_run_state(
            last_completed_index=i - 1,
            total_products=total_products,
            status="stopped"
        )
        print("Run stopped safely before starting next product.")
        sys.exit(0)

    row = products.iloc[i]

    sku = row["sku"]
    product_name = row["product_name"]
    keyword = row["image_keyword"]

    print(f"\nProcessing {i+1}/{total_products} — {product_name}")

    #########################################
    # SEARCH PRODUCT BY SKU
    #########################################

    found = search_by_sku(driver, sku)

    if not found:
        print(f"[SKU NOT FOUND IN GRAZECART] {product_name} ({sku})")

        record_skipped(
            product_name,
            sku,
            "sku_not_found_in_grazecart"
        )

        save_run_state(
            last_completed_index=i,
            total_products=total_products,
            last_completed_sku=sku,
            last_completed_product_name=product_name,
            status="running"
        )

        continue


    #########################################
    # UPLOAD PRODUCT DATA
    #########################################

    upload_product_data(row)

    if stop_requested():
        save_run_state(
            last_completed_index=i - 1,
            total_products=total_products,
            status="stopped"
        )
        print("Run stopped safely after product data upload checkpoint.")
        sys.exit(0)

    #########################################
    # GET IMAGES
    #########################################

    normalized_name = row.get("normalized_name", "")

    library_match = None

    if USE_IMAGE_LIBRARY:
        library_match = find_image_in_library(
            normalized_name=normalized_name,
            product_name=product_name,
            image_keyword=keyword,
            use_generated_library_images=USE_GENERATED_LIBRARY_IMAGES
        )

    # STEP 1 — Try UPC/GTIN candidates in Nielsen
    nielsen_match = get_nielsen_match_by_barcode(
        sku,
        keep_check_digit=KEEP_CHECK_DIGIT
    )

    # STEP 2 — Fallback to normalized_name in Nielsen
    if not nielsen_match:
        nielsen_match = get_nielsen_match(normalized_name)

    # STEP 3 — Fallback to product_name in Nielsen
    if not nielsen_match:
        nielsen_match = get_nielsen_match(product_name)

    nielsen_img = nielsen_match["image_url"] if nielsen_match else None
    pexels_img = get_pexels_image(keyword)
    #########################################
    # VALIDATE / CHOOSE IMAGE
    #########################################

    source = None
    chosen = None
    chosen_image_url_for_csv = ""

    validation_decision = ""
    best_image = ""
    image1_score = ""
    image2_score = ""
    single_image_score = ""
    single_image_decision = ""

    # CASE 0 — Image library match
    if library_match:
        print(f"Image library match found — matched on {library_match['matched_on']}")
        chosen = library_match["local_path"]
        source = library_match["source"]
        chosen_image_url_for_csv = library_match.get("image_url", "")
    # CASE 1 — Exact Nielsen match: skip validation
    if nielsen_match and nielsen_match["exact_match"]:
        print("Exact Nielsen match found — skipping validation")
        chosen = nielsen_match["image_url"]
        source = "nielsen"
        nielsen_count += 1

    # CASE 2 — Strong Nielsen score: skip validation
    elif nielsen_match and nielsen_match["score"] >= 0.82 and not pexels_img:
        print("Strong Nielsen match found — skipping validation")
        chosen = nielsen_match["image_url"]
        source = "nielsen"
        nielsen_count += 1

    # CASE 3 — Both Nielsen and Pexels exist: compare both
    elif nielsen_img and pexels_img:
        score1, score2, best, decision = validate_images(
            product_name,
            nielsen_img,
            pexels_img
        )

        image1_score = score1
        image2_score = score2
        best_image = best
        validation_decision = decision

        if decision == "USE_EXISTING":
            if best == "1":
                chosen = nielsen_img
                source = "nielsen"
                nielsen_count += 1
            elif best == "2":
                chosen = pexels_img
                source = "pexels"
                pexels_count += 1

    # CASE 4 — Nielsen only: validate single image
    elif nielsen_img:
        score, decision = validate_single_image(product_name, nielsen_img)

        single_image_score = score
        single_image_decision = decision
        validation_decision = decision

        if decision == "ACCEPT":
            chosen = nielsen_img
            source = "nielsen"
            nielsen_count += 1

    # CASE 5 — Pexels only: validate single image
    elif pexels_img:
        score, decision = validate_single_image(product_name, pexels_img)

        single_image_score = score
        single_image_decision = decision
        validation_decision = decision

        if decision == "ACCEPT":
            chosen = pexels_img
            source = "pexels"
            pexels_count += 1


    #########################################
    # LOG FOUND BUT NOT USED
    #########################################

    if not chosen:
        if nielsen_img:
            record_found_not_used(
                product_name=product_name,
                sku=sku,
                keyword=keyword,
                normalized_name=normalized_name,
                company_name=company_name,
                found_source="nielsen",
                found_image_url=nielsen_img,
                reason="found_but_not_selected",
                validation_decision=validation_decision,
                best_image=best_image,
                image1_score=image1_score,
                image2_score=image2_score,
                single_image_score=single_image_score,
                single_image_decision=single_image_decision,
                nielsen_score=nielsen_match.get("score") if nielsen_match else "",
                nielsen_exact_match=nielsen_match.get("exact_match") if nielsen_match else ""
            )

        if pexels_img:
            record_found_not_used(
                product_name=product_name,
                sku=sku,
                keyword=keyword,
                normalized_name=normalized_name,
                company_name=company_name,
                found_source="pexels",
                found_image_url=pexels_img,
                reason="found_but_not_selected",
                validation_decision=validation_decision,
                best_image=best_image,
                image1_score=image1_score,
                image2_score=image2_score,
                single_image_score=single_image_score,
                single_image_decision=single_image_decision,
                nielsen_score=nielsen_match.get("score") if nielsen_match else "",
                nielsen_exact_match=nielsen_match.get("exact_match") if nielsen_match else ""
            )
    
    #########################################
    #STOP CHECKPOINT BEFORE UPLOADING IMAGE
    #########################################

    if stop_requested():
        save_run_state(
            last_completed_index=i - 1,
            total_products=total_products,
            status="stopped"
        )
        print("Run stopped safely before image generation checkpoint.")
        sys.exit(0)

    #########################################
    # GENERATE IF NEEDED
    #########################################

    if not chosen:
        if USE_AI_IMAGE_GEN:
            print("Generating image")

            try:
                chosen = generate_image(product_name, keyword)
            except Exception as e:
                print("Image generation failed:", e)
                chosen = None

            if not chosen:
                print("Skipping image — generation failed")

                record_skipped(
                    product_name,
                    sku,
                    "image_generation_failed"
                )

                record_missing_image_review(
                    product_name=product_name,
                    sku=sku,
                    keyword=keyword,
                    normalized_name=normalized_name,
                    reason="image_generation_failed",
                    company_name=company_name,
                    ai_generation_enabled=USE_AI_IMAGE_GEN,
                    nielsen_match=nielsen_match,
                    nielsen_img=nielsen_img,
                    pexels_img=pexels_img,
                    validation_decision=validation_decision,
                    best_image=best_image,
                    image1_score=image1_score,
                    image2_score=image2_score,
                    single_image_score=single_image_score,
                    single_image_decision=single_image_decision
                )

                save_run_state(
                    last_completed_index=i,
                    total_products=total_products,
                    last_completed_sku=sku,
                    last_completed_product_name=product_name,
                    status="running"
                )

                continue

            source = "openai_generated"
            generated_count += 1

        else:
            print("Skipping image — AI generation disabled for this run")

            record_skipped(
                product_name,
                sku,
                "ai_generation_disabled_no_existing_image_selected"
            )

            record_missing_image_review(
                product_name=product_name,
                sku=sku,
                keyword=keyword,
                normalized_name=normalized_name,
                reason="ai_generation_disabled_no_existing_image_selected",
                company_name=company_name,
                ai_generation_enabled=USE_AI_IMAGE_GEN,
                nielsen_match=nielsen_match,
                nielsen_img=nielsen_img,
                pexels_img=pexels_img,
                validation_decision=validation_decision,
                best_image=best_image,
                image1_score=image1_score,
                image2_score=image2_score,
                single_image_score=single_image_score,
                single_image_decision=single_image_decision
            )

            save_run_state(
                last_completed_index=i,
                total_products=total_products,
                last_completed_sku=sku,
                last_completed_product_name=product_name,
                status="running"
            )

            continue

    #########################################
    # PREPARE IMAGE FILE
    #########################################

    if source in ["pexels", "nielsen"]:
        chosen_image_url_for_csv = chosen
        chosen = download_image(chosen, normalized_name, product_name, company_name, source)


    #########################################
    #STOP CHECKPOINT BEFORE UPLOADING IMAGE
    #########################################
    if stop_requested():
        save_run_state(
            last_completed_index=i - 1,
            total_products=total_products,
            status="stopped"
        )
        print("Run stopped safely before uploading image.")
        sys.exit(0)
    #########################################
    # UPLOAD IMAGE
    #########################################

    upload_image(chosen)

    #########################################
    # RECORD IMAGE
    #########################################

    record_image(
        product_name=product_name,
        normalized_name=normalized_name,
        image_keyword=keyword,
        sku=sku,
        source=source,
        local_path=chosen,
        image_url=chosen_image_url_for_csv
    )
    #########################################
    # SAVE STATE AFTER EACH PRODUCT
    #########################################
    save_run_state(
        last_completed_index=i,
        total_products=total_products,
        last_completed_sku=sku,
        last_completed_product_name=product_name,
        status="running"
    )

#MARK RUN AS COMPLETED
save_run_state(
    last_completed_index=total_products - 1,
    total_products=total_products,
    status="complete"
)

print("\nAll products processed.")

print("\nImage Source Summary")
print("---------------------")
print("Nielsen images:", nielsen_count)
print("Pexels images:", pexels_count)
print("Generated images:", generated_count)