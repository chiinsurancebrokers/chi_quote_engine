"""
CHI Insurance Brokers — Personalized Quote Analysis & Recommendation Narrative
Uses Claude to generate a human-readable, client-tailored explanation of WHY
a specific plan is recommended over the alternatives.

Output structure (dict):
  {
    "headline":         str   — one-line summary ("Το καλύτερο πλάνο για την οικογένειά σας")
    "main_rationale":   str   — 2-3 paragraph narrative, personalised to client profile
    "key_reasons":      list  — 3-5 bullet reasons for the recommendation
    "plan_verdicts":    list  — one verdict per plan: {"insurer", "plan", "verdict", "tag"}
    "key_concerns":     list  — 1-3 honest caveats/trade-offs the client should know
    "decision_factors": list  — the top 3 criteria that drove the recommendation
  }
"""

import json
import re
import time

import anthropic
import streamlit as st

from config import MODEL, MAX_RETRIES, RETRY_WAIT_BASE


# ─── ANALYSIS PROMPT ────────────────────────────────────────────────

def _build_analysis_prompt(
    proposals: list,
    recommended_idx: int,
    client_name: str,
    client_members: list,
) -> str:
    """Builds the prompt sent to Claude for narrative analysis."""

    rec = proposals[recommended_idx] if recommended_idx < len(proposals) else {}

    # Summarise client profile
    member_lines = []
    for m in client_members:
        member_lines.append(f"  - {m.get('role', 'Μέλος')}, ηλικία {m.get('age', '?')}")
    members_str = "\n".join(member_lines) if member_lines else "  - (δεν δόθηκαν στοιχεία)"

    # Summarise each proposal compactly
    plan_summaries = []
    for i, p in enumerate(proposals):
        tag = " ← ΠΡΟΤΕΙΝΟΜΕΝΟ" if i == recommended_idx else ""
        summary = (
            f"ΠΛΑΝΟ {i+1}{tag}: {p.get('insurer','?')} — {p.get('plan_name','?')}\n"
            f"  Ασφάλιστρο: {p.get('annual_premium','?')} {p.get('currency','EUR')} / έτος\n"
            f"  Απαλλαγή: {p.get('deductible','?')} | Μέγ. Κεφάλαιο: {p.get('max_coverage','?')}\n"
            f"  Γεωγραφία: {p.get('geography','?')} | Νοσηλεία: {p.get('hospital_class','?')}\n"
            f"  Νοσοκομειακή: {p.get('inpatient','?')}\n"
            f"  Εξωνοσ. Όριο: {p.get('outpatient_limit','?')} | Εξωνοσ. %: {p.get('outpatient_pct','?')}\n"
            f"  MRI/CT/PET: {p.get('mri_ct_pet','?')} | Καρκίνος: {p.get('cancer','?')}\n"
            f"  Χρόνιες: {p.get('chronic_conditions','?')} | Φυσιοθεραπεία: {p.get('physiotherapy','?')}\n"
            f"  Εκκένωση: {p.get('evacuation_repatriation','?')} | Οδοντ. Έκτακτη: {p.get('dental_emergency','?')}\n"
            f"  Ψυχ. Νοσηλεία: {p.get('psychiatric_inpatient','?')} | Ψυχ. Εξωτ.: {p.get('psychiatric_outpatient','?')}\n"
            f"  Αναμονή: {p.get('waiting_period','?')} | Προϋπ. Παθήσεις: {p.get('preexisting','?')}\n"
            f"  Προλ. Έλεγχος: {p.get('wellness_screening','?')} | Έλεγχος Καρκίνου: {p.get('cancer_screening','?')}\n"
            f"  Μεταμόσχευση: {p.get('organ_transplant','?')} | Ανακουφιστική: {p.get('hospice_care','?')}\n"
        )
        if p.get("key_notes"):
            summary += f"  Σημειώσεις: {'; '.join(p.get('key_notes', [])[:3])}\n"
        plan_summaries.append(summary)

    plans_block = "\n".join(plan_summaries)

    prompt = f"""
Είσαι έμπειρος ασφαλιστικός σύμβουλος της CHI Insurance Brokers.
Ο πελάτης είναι: {client_name}
Μέλη οικογένειας:
{members_str}

Παρακάτω είναι οι ασφαλιστικές προσφορές που συγκρίθηκαν:

{plans_block}

Η εταιρεία προτείνει το ΠΛΑΝΟ {recommended_idx + 1} ({rec.get('insurer','?')} — {rec.get('plan_name','?')}).

Σκοπός σου: να γράψεις μια ολοκληρωμένη, εξατομικευμένη ανάλυση που θα βοηθήσει τον πελάτη να κατανοήσει ΓΙΑΤΙ αυτό το πλάνο είναι η καλύτερη επιλογή γι' αυτόν, με βάση το προφίλ του και τα πραγματικά δεδομένα των προσφορών.

Απάντησε ΑΠΟΚΛΕΙΣΤΙΚΑ με έγκυρο JSON (χωρίς markdown backticks, χωρίς άλλο κείμενο):

{{
  "headline": "Μια σύντομη, δυναμική πρόταση που αιτιολογεί την πρόταση (π.χ. 'Η πιο ολοκληρωμένη κάλυψη για την οικογένεια {client_name} — με τη βέλτιστη σχέση κόστους/παροχών')",

  "main_rationale": "2-3 παράγραφοι (συνολικά 120-180 λέξεις). Εξήγησε συγκεκριμένα: (α) ποια χαρακτηριστικά του προτεινόμενου πλάνου ταιριάζουν στο συγκεκριμένο προφίλ του πελάτη, (β) πού υπερτερεί έναντι των άλλων, (γ) ποια οικονομική ή κλινική αξία αποκομίζει ο πελάτης. Γράψε σε β' ενικό, ζεστό επαγγελματικό ύφος, σαν να μιλάς απευθείας στον πελάτη.",

  "key_reasons": [
    "Λόγος 1: συγκεκριμένο πλεονέκτημα με αναφορά σε πραγματικά νούμερα/στοιχεία",
    "Λόγος 2: ...",
    "Λόγος 3: ...",
    "Λόγος 4 (προαιρετικός): ...",
    "Λόγος 5 (προαιρετικός): ..."
  ],

  "plan_verdicts": [
    {{
      "insurer": "Όνομα ασφαλιστικής",
      "plan": "Όνομα πλάνου",
      "verdict": "Μια πρόταση που αξιολογεί αντικειμενικά το πλάνο — τι κάνει καλά, τι του λείπει",
      "tag": "ΑΡΙΣΤΟ" ή "ΚΑΛΟ" ή "ΜΕΣΑΙΟ" ή "ΠΕΡΙΟΡΙΣΜΕΝΟ"
    }}
  ],

  "key_concerns": [
    "Ειλικρινής επισήμανση 1: κάτι που ο πελάτης πρέπει να γνωρίζει / μια αδυναμία του προτεινόμενου πλάνου",
    "Επισήμανση 2 (προαιρετική): ..."
  ],

  "decision_factors": [
    "Κριτήριο 1 που έκρινε την επιλογή",
    "Κριτήριο 2",
    "Κριτήριο 3"
  ]
}}

Κανόνες:
- Βασίσου ΑΠΟΚΛΕΙΣΤΙΚΑ στα δεδομένα που σου δόθηκαν — μην εφεύρεις στοιχεία.
- Τα key_reasons πρέπει να είναι συγκεκριμένα (π.χ. "Απαλλαγή €500 έναντι €1.000 των ανταγωνιστών") και ΌΧΙ γενικά.
- Τα plan_verdicts πρέπει να καλύπτουν ΟΛΑ τα πλάνα.
- Τα key_concerns πρέπει να είναι ειλικρινή — ο πελάτης αξίζει τη διαφάνεια.
- Η γλώσσα να είναι ελληνικά, επαγγελματική αλλά ανθρώπινη.
"""
    return prompt.strip()


# ─── MAIN ANALYSIS FUNCTION ─────────────────────────────────────────

def generate_recommendation_analysis(
    proposals: list,
    recommended_idx: int,
    client_name: str,
    client_members: list,
    api_key: str,
) -> dict:
    """
    Calls Claude to produce a personalized recommendation analysis narrative.

    Returns a dict with keys:
      headline, main_rationale, key_reasons, plan_verdicts,
      key_concerns, decision_factors
    """
    client_obj = anthropic.Anthropic(api_key=api_key)
    prompt = _build_analysis_prompt(proposals, recommended_idx, client_name, client_members)

    last_error = None
    for attempt in range(MAX_RETRIES):
        try:
            response = client_obj.messages.create(
                model=MODEL,
                max_tokens=2000,
                messages=[{"role": "user", "content": prompt}],
            )
            raw = response.content[0].text.strip()
            raw = re.sub(r"```json|```", "", raw).strip()
            return json.loads(raw)

        except anthropic.RateLimitError as e:
            last_error = e
            if attempt < MAX_RETRIES - 1:
                wait = RETRY_WAIT_BASE * (2 ** attempt)
                st.warning(
                    f"⏳ Rate limit (ανάλυση) — αναμονή {wait}s "
                    f"(απόπειρα {attempt + 1}/{MAX_RETRIES})..."
                )
                time.sleep(wait)
            else:
                raise RuntimeError(
                    f"Rate limit μετά από {MAX_RETRIES} απόπειρες."
                ) from e

        except anthropic.APIStatusError as e:
            raise RuntimeError(
                f"Claude API σφάλμα: {e.status_code} — {e.message}"
            ) from e

        except json.JSONDecodeError as e:
            raise RuntimeError(
                "Το Claude επέστρεψε μη-έγκυρο JSON στην ανάλυση. Δοκίμασε ξανά."
            ) from e

    raise RuntimeError("Αποτυχία δημιουργίας ανάλυσης.") from last_error
