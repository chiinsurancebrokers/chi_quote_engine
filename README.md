# CHI Insurance Brokers — Insurance Presentation Generator v2

Αυτόματη δημιουργία παρουσιάσεων ασφάλισης από PDF προσφορές, με χρήση Claude AI.

## Τι νέο υπάρχει στην v2

| Βελτίωση | Περιγραφή |
|---|---|
| **Modular architecture** | 4 αρχεία αντί για 1 (config / extraction / pptx_builder / app) |
| **Μηνιαίο ασφάλιστρο** | Εμφανίζεται αυτόματα σε όλες τις slides |
| **Coverage Score** | Βαθμολογία 0–10 ανά προσφορά, ορατή στο UI και στις slides |
| **8 νέα πεδία εξαγωγής** | Χρόνιες, Εκκένωση, Οδοντιατρική, Ψυχιατρική, Μεταμόσχευση, Wellness, Cancer Screening, Hospice |
| **Μοναδική Κάλυψη (★)** | Στον πίνακα σύγκρισης επισημαίνεται αυτόματα κάλυψη που έχει μόνο ένα πλάνο |
| **14 γραμμές στον πίνακα** | Από 12 → 14 με νέα πεδία |
| **Logo υποστήριξη** | Ανέβασε λογότυπο PNG/JPG από το sidebar |
| **Cache εξαγωγής** | Το ίδιο PDF δεν καλεί το API ξανά (MD5 hash cache) |
| **Key notes στο closing** | Τα βασικά πλεονεκτήματα εμφανίζονται στην τελευταία slide |
| **3-στήλη επεξεργασία** | Όλα τα πεδία διαθέσιμα στο UI (20+ πεδία ανά πρόταση) |
| **Fixes** | Null guard σε outpatient_pct, σωστός τίτλος για 2 προτάσεις, no duplicate imports |

## Δομή αρχείων

```
quote_creator/
  app.py           ← Streamlit UI μόνο
  config.py        ← Χρώματα, σταθερές, insurer_color()
  extraction.py    ← Prompt Claude + extract_insurance_data() + compute_score()
  pptx_builder.py  ← Δημιουργία PPTX (όλα τα slides)
  requirements.txt
  README.md
```

## Εγκατάσταση

```bash
# 1. Clone
git clone https://github.com/chiinsurancebrokers/quote_creator.git
cd quote_creator

# 2. Dependencies
pip install -r requirements.txt

# 3. Εκκίνηση
streamlit run app.py
```

## Χρήση

1. Άνοιξε **http://localhost:8501**
2. (Προαιρετικά) Ανέβασε λογότυπο PNG στο sidebar
3. Βάλε το **Claude API key** (από https://console.anthropic.com)
4. Συμπλήρωσε στοιχεία πελάτη (όνομα, ηλικίες μελών)
5. Ανέβασε **2–4 PDF** προσφορές
6. Πάτα **«Ανάλυση με Claude API»**
7. Έλεγξε / επεξεργάσου τα εξαχθέντα στοιχεία στις tabs
8. Επέλεξε ποια πρόταση είναι «Προτεινόμενη»
9. Πάτα **«Δημιουργία PPTX»** → Download!

## Deploy στο Streamlit Cloud (δωρεάν)

1. Ανέβασε τον κώδικα στο GitHub
2. Πήγαινε στο https://share.streamlit.io
3. Σύνδεσε το GitHub repo, επέλεξε `app.py`
4. Στα **Secrets** πρόσθεσε: `Claude_API_Key = "sk-ant-..."`
5. Deploy!

## Supported Insurers (brand colors)

AXA · Generali · Morgan Price · NOW Health · ERGO · Allianz · Cigna

## Στοιχεία

- **Broker:** CHI Insurance Brokers
- **Email:** info@chiinsurancebrokers.com
- **Tel:** +30 697 590 0189
