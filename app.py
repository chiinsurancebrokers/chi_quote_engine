"""
CHI Insurance Brokers — Αυτόματη Δημιουργία Παρουσίασης
Εκκίνηση: streamlit run app.py
"""

import hashlib
import time
from datetime import datetime

import streamlit as st

from config import BROKER_DEFAULTS, INTER_FILE_DELAY
from extraction import compute_score, extract_insurance_data
from pptx_builder import generate_pptx


st.set_page_config(
    page_title="CHI Insurance — Παρουσιάσεις",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)


def main():
    # ── GLOBAL STYLES ────────────────────────────────────────────────
    st.markdown("""
    <style>
    .main { background: #F4F9FF; }
    .stButton > button {
        background: #1C3F5E; color: white; border-radius: 8px;
        font-weight: bold; padding: 0.6em 2em; border: none;
    }
    .stButton > button:hover { background: #00B4D8; }
    div[data-testid="stFileUploader"] {
        border: 2px dashed #00B4D8; border-radius: 8px; padding: 1em;
    }
    </style>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns([1, 5])
    with c1:
        st.markdown("# 🛡️")
    with c2:
        st.markdown("## CHI Insurance Brokers")
        st.markdown("*Αυτόματη Δημιουργία Παρουσιάσεων Ασφάλισης*")
    st.divider()

    # ── SIDEBAR ──────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⚙️ Ρυθμίσεις")

        # API key — prefer secrets, fall back to text input
        _secret_key = (
            st.secrets.get("Claude_API_Key")
            or st.secrets.get("ANTHROPIC_API_KEY")
            or st.secrets.get("claude_api_key")
            or ""
        )
        if _secret_key:
            api_key = _secret_key
            st.success("🔑 API Key φορτώθηκε αυτόματα", icon="✅")
        else:
            api_key = st.text_input(
                "🔑 Claude API Key", type="password",
                help="Ή πρόσθεσέ το στο Secrets ως: Claude_API_Key = 'sk-ant-...'"
            )

        st.markdown("---")
        st.markdown("### 🖼️ Λογότυπο")
        logo_file  = st.file_uploader(
            "Ανέβασε λογότυπο (PNG / JPG)",
            type=["png", "jpg", "jpeg"],
            help="Προαιρετικό. Εμφανίζεται στην κεντρική σελίδα της παρουσίασης."
        )
        logo_bytes = logo_file.read() if logo_file else None

        st.markdown("---")
        st.markdown("### 👤 Στοιχεία Μεσίτη")
        broker_name  = st.text_input("Όνομα",    value=BROKER_DEFAULTS["name"])
        broker_tel   = st.text_input("Τηλέφωνο", value=BROKER_DEFAULTS["tel"])
        broker_email = st.text_input("Email",     value=BROKER_DEFAULTS["email"])

        st.markdown("---")
        st.markdown("### 👥 Στοιχεία Πελάτη")
        client_name = st.text_input(
            "Επώνυμο / Όνομα Πελάτη",
            placeholder="π.χ. Τοτικίδη Κατία"
        )

        st.markdown("**Μέλη:**")
        n_members = st.number_input("Αριθμός μελών", 1, 6, 2)
        members = []
        for i in range(n_members):
            mc1, mc2 = st.columns(2)
            with mc1:
                age = st.number_input(
                    f"Ηλικία #{i + 1}", 0, 99,
                    30 if i == 0 else 17, key=f"age_{i}"
                )
            with mc2:
                role = st.selectbox(
                    "Ρόλος",
                    ["Κύρια Ασφαλισμένη", "Κύριος Ασφαλισμένος",
                     "Εξαρτώμενο Μέλος", "Σύζυγος"],
                    key=f"role_{i}"
                )
            members.append({"age": age, "role": role})

    # ── PDF UPLOAD ───────────────────────────────────────────────────
    st.markdown("### 📄 Φόρτωσε τις Ασφαλιστικές Προσφορές (PDF)")
    st.info("Φόρτωσε 2–4 PDF προσφορές. Το Claude θα εξάγει αυτόματα όλα τα στοιχεία.", icon="ℹ️")

    uploaded_files = st.file_uploader(
        "Επίλεξε PDF αρχεία", type="pdf", accept_multiple_files=True,
        help="Ανέβασε τις προσφορές Generali, Morgan Price, NOW Health κ.λπ."
    )

    if not uploaded_files:
        st.markdown("---")
        st.markdown("#### Πώς λειτουργεί:")
        h1, h2, h3 = st.columns(3)
        with h1:
            st.markdown("**1️⃣ Ανέβασε PDFs**\nΌλες οι προσφορές που θέλεις να συγκρίνεις")
        with h2:
            st.markdown("**2️⃣ Claude τα αναλύει**\nΕξάγει αυτόματα κεφάλαια, απαλλαγές, καλύψεις")
        with h3:
            st.markdown("**3️⃣ Download PPTX**\nΈτοιμη παρουσίαση με το brand σου")
        return

    # ── SESSION STATE INIT ───────────────────────────────────────────
    if "proposals"  not in st.session_state: st.session_state.proposals  = {}
    if "pdf_cache"  not in st.session_state: st.session_state.pdf_cache  = {}

    # ── EXTRACTION ───────────────────────────────────────────────────
    if st.button("🤖 Ανάλυση με Claude API", type="primary", disabled=not api_key):
        if not api_key:
            st.error("Χρειάζεσαι Claude API key!")
            return

        progress = st.progress(0, text="Αρχικοποίηση...")
        st.session_state.proposals = {}
        total = len(uploaded_files)

        for idx, uf in enumerate(uploaded_files):
            progress.progress(idx / total, text=f"Ανάλυση {idx + 1}/{total}: {uf.name}...")
            try:
                pdf_bytes = uf.read()
                pdf_hash  = hashlib.md5(pdf_bytes).hexdigest()

                if pdf_hash in st.session_state.pdf_cache:
                    # Return cached result — no API call needed
                    data = st.session_state.pdf_cache[pdf_hash]
                    st.success(f"⚡ {uf.name} — φορτώθηκε από cache")
                else:
                    data = extract_insurance_data(pdf_bytes, api_key, filename=uf.name)
                    st.session_state.pdf_cache[pdf_hash] = data
                    st.success(
                        f"✅ {uf.name} → {data.get('insurer', '')} {data.get('plan_name', '')}"
                    )

                st.session_state.proposals[uf.name] = data

            except Exception as e:
                st.error(f"❌ Σφάλμα στο {uf.name}: {e}")

            if idx < total - 1:
                time.sleep(INTER_FILE_DELAY)

        progress.progress(1.0, text="✅ Ολοκληρώθηκε!")

    # ── DISPLAY & EDIT EXTRACTED DATA ───────────────────────────────
    if st.session_state.get("proposals"):
        proposals_list = list(st.session_state.proposals.values())
        file_names     = list(st.session_state.proposals.keys())

        st.markdown("---")
        st.markdown("### 📊 Εξαχθέντα Στοιχεία")

        # Βαθμολογία κάλυψης
        st.markdown("#### Βαθμολογία Κάλυψης")
        score_cols = st.columns(len(proposals_list))
        for col, prop in zip(score_cols, proposals_list):
            sc    = compute_score(prop)
            emoji = "🟢" if sc >= 7 else ("🟡" if sc >= 5 else "🔴")
            with col:
                st.metric(
                    label=f"{prop.get('insurer','?')} — {prop.get('plan_name','?')[:18]}",
                    value=f"{sc} / 10",
                    delta=f"{emoji} Βαθμολογία Κάλυψης",
                )

        st.caption("Μπορείς να επεξεργαστείς οποιοδήποτε πεδίο πριν τη δημιουργία.")
        st.markdown("---")

        # Editable tabs — one per proposal
        edited_proposals = []
        tabs = st.tabs([
            f"📋 {p.get('insurer', '?')} — {p.get('plan_name', '?')[:20]}"
            for p in proposals_list
        ])

        for tab, prop, fname in zip(tabs, proposals_list, file_names):
            with tab:
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.markdown("**📌 Βασικά Στοιχεία**")
                    prop["insurer"]        = st.text_input("Ασφαλιστική",       prop.get("insurer", ""),                  key=f"ins_{fname}")
                    prop["plan_name"]      = st.text_input("Πλάνο",             prop.get("plan_name", ""),                key=f"plan_{fname}")
                    prop["annual_premium"] = st.text_input("Ετήσιο Ασφάλιστρο", str(prop.get("annual_premium", "")),      key=f"prem_{fname}")
                    prop["currency"]       = st.selectbox("Νόμισμα",            ["EUR", "USD", "GBP"],
                                                          index=["EUR","USD","GBP"].index(prop.get("currency","EUR") or "EUR"),
                                                          key=f"cur_{fname}")
                    prop["deductible"]     = st.text_input("Απαλλαγή",          prop.get("deductible", ""),               key=f"ded_{fname}")
                    prop["max_coverage"]   = st.text_input("Μέγιστο Κεφάλαιο",  str(prop.get("max_coverage", "")),        key=f"maxcov_{fname}")
                    prop["geography"]      = st.text_input("Γεωγραφία",          prop.get("geography", ""),                key=f"geo_{fname}")
                    prop["hospital_class"] = st.text_input("Θέση Νοσηλείας",    prop.get("hospital_class", ""),           key=f"hosp_{fname}")
                    prop["waiting_period"] = st.text_input("Αναμονή",            prop.get("waiting_period", ""),           key=f"wait_{fname}")
                    prop["preexisting"]    = st.text_input("Προϋπ. Παθήσεις",   prop.get("preexisting", ""),              key=f"preex_{fname}")

                with col2:
                    st.markdown("**✅ Καλύψεις**")
                    prop["inpatient"]               = st.text_input("Νοσηλεία",              prop.get("inpatient", ""),               key=f"inp_{fname}")
                    prop["outpatient_limit"]         = st.text_input("Εξωνοσοκ. Όριο",       str(prop.get("outpatient_limit", "")),   key=f"outp_{fname}")
                    prop["outpatient_pct"]           = st.text_input("Εξωνοσοκ. %",           str(prop.get("outpatient_pct") or ""),   key=f"outpct_{fname}")
                    prop["mri_ct_pet"]               = st.text_input("MRI / CT / PET",        prop.get("mri_ct_pet", ""),              key=f"mri_{fname}")
                    prop["cancer"]                   = st.text_input("Καρκίνος",               prop.get("cancer", ""),                  key=f"can_{fname}")
                    prop["physiotherapy"]            = st.text_input("Φυσιοθεραπεία",          prop.get("physiotherapy", ""),           key=f"physio_{fname}")
                    prop["chronic_conditions"]       = st.text_input("Χρόνιες Παθήσεις",      prop.get("chronic_conditions", ""),      key=f"chron_{fname}")
                    prop["evacuation_repatriation"]  = st.text_input("Εκκένωση / Μεταφορά",  prop.get("evacuation_repatriation", ""), key=f"evac_{fname}")
                    prop["psychiatric_inpatient"]    = st.text_input("Ψυχ. Νοσηλεία",         prop.get("psychiatric_inpatient", ""),   key=f"psyin_{fname}")
                    prop["psychiatric_outpatient"]   = st.text_input("Ψυχ. Εξωτερικά",       prop.get("psychiatric_outpatient", ""),  key=f"psyout_{fname}")

                with col3:
                    st.markdown("**➕ Πρόσθετα & Παρατηρήσεις**")
                    prop["dental_emergency"]   = st.text_input("Οδοντ. Έκτακτη",         prop.get("dental_emergency", ""),   key=f"dent_{fname}")
                    prop["wellness_screening"] = st.text_input("Προληπτικός Έλεγχος",     prop.get("wellness_screening", ""), key=f"well_{fname}")
                    prop["cancer_screening"]   = st.text_input("Έλεγχος Καρκίνου",        prop.get("cancer_screening", ""),   key=f"canscr_{fname}")
                    prop["organ_transplant"]   = st.text_input("Μεταμόσχευση Οργάνου",    prop.get("organ_transplant", ""),   key=f"organ_{fname}")
                    prop["hospice_care"]       = st.text_input("Ανακουφιστική Φροντίδα",  prop.get("hospice_care", ""),       key=f"hosp2_{fname}")
                    prop["home_nursing"]       = st.text_input("Νοσηλεία Κατ' Οίκον",    prop.get("home_nursing", ""),       key=f"homenur_{fname}")

                    st.markdown("**💳 Τρόπος Πληρωμής**")
                    freq_options = ["Μηνιαία", "Τριμηνιαία", "Εξαμηνιαία", "Ετήσια"]
                    current_freq = prop.get("payment_frequency") or "Ετήσια"
                    if current_freq not in freq_options:
                        current_freq = "Ετήσια"
                    prop["payment_frequency"] = st.selectbox(
                        "Συχνότητα πληρωμής",
                        freq_options,
                        index=freq_options.index(current_freq),
                        key=f"freq_{fname}",
                        help="Επιλέξτε πώς θα εμφανίζεται το ασφάλιστρο στην παρουσίαση. Το ετήσιο κεφάλαιο διαιρείται αυτόματα."
                    )

                    st.markdown("**📝 Παρατηρήσεις**")
                    notes_raw  = prop.get("key_notes") or []
                    notes_str  = "\n".join(notes_raw) if isinstance(notes_raw, list) else str(notes_raw)
                    edited_notes = st.text_area(
                        "Μία παρατήρηση ανά γραμμή",
                        notes_str, height=150, key=f"notes_{fname}"
                    )
                    prop["key_notes"] = [
                        n.strip() for n in edited_notes.splitlines() if n.strip()
                    ]

                edited_proposals.append(prop)

        # ── RECOMMENDED CHOICE ───────────────────────────────────────
        st.markdown("---")
        st.markdown("### 🎯 Επιλογή Πρότασης")
        insurer_labels = [
            f"{p.get('insurer', '')} — {p.get('plan_name', '')} "
            f"({p.get('currency','€')}{p.get('annual_premium', '—')})"
            for p in edited_proposals
        ]
        rec_idx = st.selectbox(
            "Ποια πρόταση να εμφανίζεται ως **ΠΡΟΤΕΙΝΟΜΕΝΗ**;",
            range(len(insurer_labels)),
            format_func=lambda i: insurer_labels[i],
        )

        # ── GENERATE ─────────────────────────────────────────────────
        st.markdown("---")
        if st.button("🎨 Δημιουργία Παρουσίασης PPTX", type="primary"):
            if not client_name:
                st.warning("Συμπλήρωσε το όνομα του πελάτη στο sidebar!")
                return

            with st.spinner("Δημιουργία παρουσίασης..."):
                try:
                    pptx_bytes = generate_pptx(
                        client_name=client_name,
                        client_members=members,
                        proposals=edited_proposals,
                        recommended_idx=rec_idx,
                        broker_name=broker_name,
                        broker_tel=broker_tel,
                        broker_email=broker_email,
                        logo_bytes=logo_bytes,
                    )
                    fname_out = (
                        f"{client_name.replace(' ', '_')}_Insurance_"
                        f"{datetime.now().strftime('%Y%m')}.pptx"
                    )
                    st.download_button(
                        label="⬇️ Download Παρουσίαση",
                        data=pptx_bytes,
                        file_name=fname_out,
                        mime=(
                            "application/vnd.openxmlformats-officedocument"
                            ".presentationml.presentation"
                        ),
                    )
                    st.success(f"✅ Η παρουσίαση '{fname_out}' είναι έτοιμη!")

                except Exception as e:
                    st.error(f"Σφάλμα: {e}")
                    import traceback
                    st.code(traceback.format_exc())

    elif uploaded_files and not st.session_state.get("proposals"):
        st.info("👆 Πάτα 'Ανάλυση με Claude API' για να εξαχθούν τα στοιχεία από τα PDFs.")


if __name__ == "__main__":
    main()
