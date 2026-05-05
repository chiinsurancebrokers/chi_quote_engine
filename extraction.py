"""
CHI Insurance Brokers — PDF Data Extraction & Coverage Scoring
Uses Claude to parse insurance PDFs into structured JSON,
then computes a weighted coverage quality score (0–10).

Smart extraction:
  - Αν το PDF έχει > 6 σελίδες ή > 60 KB, εξάγει κείμενο με PyMuPDF
    και κρατά μόνο τις σελίδες με ασφαλιστικά δεδομένα (scoring).
  - Στέλνει κείμενο αντί για binary PDF → πολύ λιγότερα tokens.
  - Fallback σε binary PDF για μικρά/απλά αρχεία.
"""

import base64
import json
import re
import time

import anthropic
import streamlit as st

try:
    import fitz  # PyMuPDF
    _PYMUPDF_OK = True
except ImportError:
    _PYMUPDF_OK = False

from config import MODEL, MAX_RETRIES, RETRY_WAIT_BASE


# ─── ΣΤΑΘΕΡΕΣ SMART EXTRACTION ──────────────────────────────────────

# Λέξεις-κλειδιά που δείχνουν χρήσιμη ασφαλιστική σελίδα
_HIGH_VALUE = [
    "ασφάλιστρο", "ασφαλίστρου", "ασφαλίστρων", "ασφαλίστρα",
    "κάλυψη", "κάλυψης", "καλύψεις", "καλύπτεται", "καλύπτονται",
    "παροχές", "παροχή", "παροχών",
    "απαλλαγή", "απαλλαγής",
    "νοσηλεία", "νοσηλείας",
    "ετήσιο", "ετήσιος", "ετήσια",
    "ανώτατο όριο", "ανώτατο",
    "χημειοθεραπεία", "ακτινοθεραπεία",
    "εξωνοσοκομειακ", "εξωτερικ",
    "διαγνωστικ", "φυσιοθεραπεί",
    "premium", "deductible", "coverage",
    "πλάνο", "πρόγραμμα",
    # Πεδία τιμολόγησης / ασφαλίστρου
    "ανάλυση ασφαλίστρων", "βασικά στοιχεία προγράμματος",
    "συνολικό καθαρό", "σύνολο δόσης", "σύνολο πρώτης δόσης",
    "καθαρό ασφάλιστρο", "ετήσιο καθαρό", "δικαίωμα",
    "συχνότητα πληρωμής", "τρόπος πληρωμής",
    "full health", "full επείγοντα",
]

# Λέξεις που μειώνουν αξία σελίδας
_LOW_VALUE = [
    "INTERNAL",
    "εναλλακτικής επίλυσης διαφορών",
    "φερεγγυότητα",
    "ν. 4364", "ν. 2496",
    "νόμος 4364",
    "εναντίωσ",
    "υπαναχώρησ",
    "τράπεζα της ελλάδος",
    "ερωτηματολόγιο αναγκών",   # Τελευταία σελίδα με ναι/όχι
    # Νομικές σελίδες 3-5 από προσυμβατική (5ψήφιες σελίδες)
    "σελίδα 3 από 5",
    "σελίδα 4 από 5",
    "σελίδα 5 από 5",
    "σελίδα 3 από 4",
    "σελίδα 4 από 4",
]

# Όρια για smart extraction
_SIZE_THRESHOLD_BYTES = 60_000   # > 60 KB → χρήση text extraction
_PAGE_THRESHOLD       = 6        # > 6 σελίδες → χρήση text extraction
_MAX_CHARS_TO_CLAUDE  = 16_000   # Μέγιστοι χαρακτήρες προς Claude


# ─── ΒΑΘΜΟΛΟΓΗΣΗ ΣΕΛΙΔΑΣ ────────────────────────────────────────────

def _score_page(text: str) -> int:
    """Βαθμολογεί μια σελίδα ως προς τη χρησιμότητά της (0–100)."""
    t = text.lower()
    score = 0
    for kw in _HIGH_VALUE:
        if kw.lower() in t:
            score += 5
    for kw in _LOW_VALUE:
        if kw.lower() in t:
            score -= 15
    return max(0, score)


# ─── SMART PDF → TEXT ────────────────────────────────────────────────

def smart_pdf_to_text(pdf_bytes: bytes, filename: str = "") -> str | None:
    """
    Εξάγει κείμενο από τις πιο σχετικές σελίδες του PDF.

    Επιστρέφει:
      - str  : το συμπιεσμένο κείμενο αν η εξαγωγή πέτυχε
      - None : αν πρέπει να σταλεί ολόκληρο το PDF binary
    """
    if not _PYMUPDF_OK:
        return None

    needs_smart = (
        len(pdf_bytes) > _SIZE_THRESHOLD_BYTES
        or _quick_page_count(pdf_bytes) > _PAGE_THRESHOLD
    )
    if not needs_smart:
        return None   # Μικρό PDF — στείλε binary κανονικά

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        total = len(doc)

        # Βαθμολόγησε κάθε σελίδα
        scored = []
        for i, page in enumerate(doc):
            text = page.get_text().strip()
            sc   = _score_page(text)
            scored.append((i + 1, sc, text))   # (αρ. σελίδας, score, κείμενο)

        doc.close()

        # ── Επιλογή σελίδων ──
        # 1. Πάντα κράτα τις πρώτες 2 σελίδες (εξώφυλλο + βασικά)
        must_keep = {1, 2}
        selected_pages = []
        total_chars    = 0

        # Πρώτα must_keep (με σειρά)
        for pg, sc, tx in scored:
            if pg in must_keep and tx:
                selected_pages.append((pg, tx))
                total_chars += len(tx)

        # Μετά υπόλοιπες με score > 0 (ταξινομημένες κατά score φθίνον)
        rest = [(pg, sc, tx) for pg, sc, tx in scored if pg not in must_keep and sc > 0]
        rest.sort(key=lambda x: -x[1])

        for pg, sc, tx in rest:
            if total_chars + len(tx) > _MAX_CHARS_TO_CLAUDE:
                continue   # Παράλειψε αν είναι πολύ μεγάλη
            selected_pages.append((pg, tx))
            total_chars += len(tx)

        # Τελική ταξινόμηση κατά αριθμό σελίδας
        selected_pages.sort(key=lambda x: x[0])

        if not selected_pages:
            return None   # Δεν εξαχθηκε τίποτα χρήσιμο

        kept    = [pg for pg, _ in selected_pages]
        skipped = total - len(kept)

        header = (
            f"=== ΑΣΦΑΛΙΣΤΙΚΗ ΠΡΟΣΦΟΡΑ: {filename} ===\n"
            f"[Εξαγωγή {len(kept)}/{total} σελίδων — {skipped} νομικές σελίδες παραλείφθηκαν]\n\n"
        )
        body = "\n\n---\n\n".join(
            f"[ΣΕΛΙΔΑ {pg}]\n{tx}" for pg, tx in selected_pages
        )

        return header + body

    except Exception:
        return None   # Fallback σε binary


def _quick_page_count(pdf_bytes: bytes) -> int:
    """Μετράει γρήγορα τις σελίδες χωρίς πλήρες parse."""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        n   = len(doc)
        doc.close()
        return n
    except Exception:
        return 0


# ─── EXTRACTION PROMPT ──────────────────────────────────────────────
# FIX: insured_members is now an empty array — no hardcoded example ages.
# Claude will populate it only if the PDF contains member information.

EXTRACT_PROMPT = """
Διάβασε αυτή την ασφαλιστική προσφορά και εξάγαγε τα παρακάτω στοιχεία σε JSON.
Απάντησε ΜΟΝΟ με valid JSON, χωρίς markdown backticks ή οποιοδήποτε άλλο κείμενο.

{
  "insurer": "Όνομα ασφαλιστικής (π.χ. Generali, Morgan Price, NOW Health, ERGO, AXA)",
  "plan_name": "Ακριβές όνομα πλάνου (π.χ. Evolution Standard, Foundation)",
  "annual_premium": "Ετήσιο ασφάλιστρο — μόνο αριθμός χωρίς σύμβολο (π.χ. 2626)",
  "currency": "EUR ή USD ή GBP",
  "deductible": "Απαλλαγή (π.χ. 500 ή 1000 ανά άτομο)",
  "max_coverage": "Μέγιστο κεφάλαιο — μόνο αριθμός (π.χ. 500000)",
  "geography": "Γεωγραφική κάλυψη (π.χ. Ευρώπη, Παγκόσμια εκτός ΗΠΑ)",
  "hospital_class": "Θέση νοσηλείας (π.χ. Α, Β, Standard Room)",

  "inpatient": "Full Refund ή ποσοστό ή Not Covered",
  "outpatient_limit": "Όριο εξωνοσοκομειακών — αριθμός ή Not Covered",
  "outpatient_pct": "Ποσοστό κάλυψης εξωνοσοκομειακών — αριθμός ή null",
  "mri_ct_pet": "Full Refund ή Not Covered ή σύντομη περιγραφή",
  "cancer": "Full Refund ή Not Covered ή σύντομη περιγραφή",
  "physiotherapy": "Ποσό ή Full Refund ή Not Covered",
  "chronic_conditions": "Full Refund ή Not Covered ή περιγραφή",
  "evacuation_repatriation": "Full Refund ή Not Covered ή περιγραφή",
  "dental_emergency": "Full Refund ή Not Covered ή ποσό",
  "wellness_screening": "Ποσό ή Not Covered (π.χ. 300)",
  "cancer_screening": "Ποσό ή Not Covered (π.χ. 1000)",
  "organ_transplant": "Ποσό ή Full Refund ή Not Covered",
  "hospice_care": "Full Refund ή Not Covered ή περιγραφή",
  "psychiatric_inpatient": "Περιγραφή ή Not Covered (π.χ. 100 days/lifetime)",
  "psychiatric_outpatient": "Περιγραφή ή Not Covered",
  "home_nursing": "Περιγραφή ή Not Covered",

  "waiting_period": "Αναμονή για παθήσεις (π.χ. Άμεση ή 6 μήνες ή 24 μήνες)",
  "preexisting": "Κάλυψη προϋπαρχουσών παθήσεων (π.χ. Άμεση MHD ή μετά 12 μήνες ή Όχι)",
  "payment_frequency": "Συχνότητα πληρωμής ασφαλίστρου: Μηνιαία ή Τριμηνιαία ή Εξαμηνιαία ή Ετήσια",

  "insured_members": [],

  "key_notes": ["σύντομη παρατήρηση 1", "σύντομη παρατήρηση 2"]
}

Κανόνες:
- Αν κάποιο πεδίο δεν βρεθεί στο PDF, βάλε null.
- Για insured_members: αν βρεις ασφαλισμένα άτομα στο PDF, συμπλήρωσε {"age": N, "role": "..."},
  αλλιώς άφησε κενό array [].
- Για key_notes: γράψε 2–4 σύντομες, σημαντικές παρατηρήσεις μόνο αν υπάρχουν στο κείμενο.
- Μην εφεύρεις πληροφορίες που δεν υπάρχουν στο PDF.
"""


# ─── PDF EXTRACTION ─────────────────────────────────────────────────

def extract_insurance_data(pdf_bytes: bytes, api_key: str, filename: str = "") -> dict:
    """
    Εξάγει ασφαλιστικά δεδομένα από PDF με Claude.

    Για μεγάλα PDFs (> 6 σελίδες ή > 60 KB): εξάγει έξυπνα μόνο
    τις σχετικές σελίδες ως κείμενο — πολύ λιγότερα tokens.
    Για μικρά PDFs: στέλνει το binary απευθείας.
    """
    client = anthropic.Anthropic(api_key=api_key)

    # ── Απόφαση: text extraction ή binary PDF ──
    extracted_text = smart_pdf_to_text(pdf_bytes, filename)

    if extracted_text:
        # Έξυπνη εξαγωγή — στέλνει μόνο το κείμενο
        pages_kept = extracted_text.count("[ΣΕΛΙΔΑ ")
        st.info(
            f"📄 «{filename}»: μεγάλο PDF — "
            f"εξαγωγή {pages_kept} σελίδων (νομικά κείμενα παραλείφθηκαν)",
            icon="✂️"
        )
        user_content = [
            {"type": "text", "text": extracted_text},
            {"type": "text", "text": EXTRACT_PROMPT},
        ]
    else:
        # Μικρό PDF — στείλε binary
        pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")
        user_content = [
            {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": pdf_b64,
                },
            },
            {"type": "text", "text": EXTRACT_PROMPT},
        ]

    last_error = None
    for attempt in range(MAX_RETRIES):
        try:
            response = client.messages.create(
                model=MODEL,
                max_tokens=2000,
                messages=[{
                    "role": "user",
                    "content": user_content,
                }],
            )
            raw = response.content[0].text.strip()
            raw = re.sub(r"```json|```", "", raw).strip()
            return json.loads(raw)

        except anthropic.RateLimitError as e:
            last_error = e
            if attempt < MAX_RETRIES - 1:
                wait  = RETRY_WAIT_BASE * (2 ** attempt)   # 10s → 20s → 40s
                label = f" ({filename})" if filename else ""
                st.warning(
                    f"⏳ Rate limit{label} — αναμονή {wait}s "
                    f"(απόπειρα {attempt + 1}/{MAX_RETRIES})..."
                )
                time.sleep(wait)
            else:
                raise RuntimeError(
                    f"Rate limit μετά από {MAX_RETRIES} απόπειρες. "
                    "Δοκίμασε ξανά σε λίγο ή μείωσε τον αριθμό PDFs."
                ) from e

        except anthropic.APIStatusError as e:
            raise RuntimeError(
                f"Claude API σφάλμα: {e.status_code} — {e.message}"
            ) from e

        except json.JSONDecodeError as e:
            raise RuntimeError(
                "Το Claude επέστρεψε μη-έγκυρο JSON. "
                "Δοκίμασε ξανά ή έλεγξε το PDF."
            ) from e

    raise RuntimeError("Αποτυχία εξαγωγής δεδομένων.") from last_error


# ─── COVERAGE SCORE ─────────────────────────────────────────────────

def compute_score(prop: dict) -> float:
    """
    Return a weighted coverage quality score from 0.0 to 10.0.

    Weights reflect clinical and financial importance of each benefit.
    Higher score = broader, richer coverage.
    """

    def covered(field: str) -> bool:
        v = prop.get(field)
        if v is None:
            return False
        s = str(v).strip()
        return s not in ("", "null", "None", "—") and "Not Covered" not in s

    weights = {
        # Core medical coverage (heavy weight)
        "inpatient":              20,
        "cancer":                 15,
        "mri_ct_pet":              8,
        "chronic_conditions":      8,
        "evacuation_repatriation": 8,
        # Outpatient (numeric — scored below)
        "outpatient_limit":        8,
        # Max coverage ceiling (numeric — scored below)
        "max_coverage":            7,
        # Secondary benefits
        "physiotherapy":           5,
        "psychiatric_outpatient":  5,
        "dental_emergency":        5,
        "wellness_screening":      3,
        "cancer_screening":        3,
        "organ_transplant":        3,
        "hospice_care":            2,
    }

    score = 0.0

    # Binary fields (covered = full weight, else 0)
    binary = [
        "inpatient", "cancer", "mri_ct_pet", "chronic_conditions",
        "evacuation_repatriation", "physiotherapy", "psychiatric_outpatient",
        "dental_emergency", "wellness_screening", "cancer_screening",
        "organ_transplant", "hospice_care",
    ]
    for field in binary:
        if covered(field):
            score += weights[field]

    # Outpatient limit — graded by amount
    ol = prop.get("outpatient_limit")
    if ol and str(ol).strip() not in ("Not Covered", "null", "None", ""):
        try:
            v = int(str(ol).replace(",", ""))
            if   v >= 5_000: score += weights["outpatient_limit"]
            elif v >= 2_000: score += weights["outpatient_limit"] * 0.7
            elif v >  0:     score += weights["outpatient_limit"] * 0.4
        except (ValueError, TypeError):
            if covered("outpatient_limit"):
                score += weights["outpatient_limit"] * 0.5

    # Max coverage ceiling — graded by amount
    mc = prop.get("max_coverage")
    if mc:
        try:
            v = int(str(mc).replace(",", ""))
            if   v >= 1_000_000: score += weights["max_coverage"]
            elif v >=   500_000: score += weights["max_coverage"] * 0.8
            elif v >=   250_000: score += weights["max_coverage"] * 0.5
        except (ValueError, TypeError):
            pass

    total_weight = sum(weights.values())   # 100 total
    return round((score / total_weight) * 10, 1)
