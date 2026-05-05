"""
CHI Insurance Brokers — Shared Configuration
Colors, model settings, broker defaults, and utility helpers.
"""

from pptx.dml.color import RGBColor


# ─── COLOR HELPERS ──────────────────────────────────────────────────

def rgb(r, g, b) -> RGBColor:
    return RGBColor(r, g, b)


C = {
    "navy":      RGBColor(0x1C, 0x3F, 0x5E),
    "navyDark":  RGBColor(0x0F, 0x26, 0x38),
    "teal":      RGBColor(0x00, 0xB4, 0xD8),
    "white":     RGBColor(0xFF, 0xFF, 0xFF),
    "offWhite":  RGBColor(0xF4, 0xF9, 0xFF),
    "textDark":  RGBColor(0x1A, 0x2B, 0x3C),
    "green":     RGBColor(0x27, 0xAE, 0x60),
    "orange":    RGBColor(0xE6, 0x7E, 0x22),
    "red":       RGBColor(0xE7, 0x4C, 0x3C),
    "gold":      RGBColor(0xF5, 0x9E, 0x0B),
    "generali":  RGBColor(0xCC, 0x00, 0x00),
    "now":       RGBColor(0x7B, 0x2D, 0x8B),
    "blue":      RGBColor(0x3B, 0x82, 0xF6),
    "axa":       RGBColor(0x00, 0x00, 0x8B),
    "allianz":   RGBColor(0x00, 0x67, 0xB1),
    "cigna":     RGBColor(0x00, 0x61, 0xA0),
    "ethniki":   RGBColor(0x00, 0x5B, 0xAA),   # Εθνική Ασφαλιστική — μπλε
    "interlife": RGBColor(0xE8, 0x00, 0x00),   # Interlife — κόκκινο
    "eurolife":  RGBColor(0x00, 0x40, 0x80),   # Eurolife FFH — σκούρο μπλε
    "groupama":  RGBColor(0x00, 0x82, 0x40),   # Groupama — πράσινο
}


def insurer_color(name: str) -> RGBColor:
    """Return a brand color for a known insurer, or teal as default."""
    n = name.upper()
    if "GENERALI"   in n: return C["generali"]
    if "MORGAN"     in n: return C["navy"]
    if "NOW"        in n: return C["now"]
    if "ERGO"       in n: return rgb(0x00, 0x5A, 0xA0)
    if "AXA"        in n: return C["axa"]
    if "ALLIANZ"    in n: return C["allianz"]
    if "CIGNA"      in n: return C["cigna"]
    if "ΕΘΝΙΚ"      in n: return C["ethniki"]   # Εθνική Ασφαλιστική
    if "ETHNIKI"    in n: return C["ethniki"]
    if "INTERLIFE"  in n: return C["interlife"]
    if "EUROLIFE"   in n: return C["eurolife"]
    if "GROUPAMA"   in n: return C["groupama"]
    return C["teal"]


# ─── API / RATE-LIMIT SETTINGS ──────────────────────────────────────
# Τα παρακάτω είναι σχεδιασμένα για Tier-1 API keys (χαμηλά rate limits).
# Αύξησε το TIER εάν έχεις υψηλότερο tier — θα μειωθούν οι αναμονές.

API_TIER = 1   # Άλλαξε σε 2 ή 3 αν έχεις υψηλότερο tier

MODEL           = "claude-sonnet-4-20250514"

# Tier 1: 5 απόπειρες × (30s, 60s, 120s, 240s, 480s)
# Tier 2: 4 απόπειρες × (15s, 30s, 60s, 120s)
# Tier 3: 3 απόπειρες × (10s, 20s, 40s)
_TIER_CONFIG = {
    1: {"max_retries": 5, "wait_base": 30, "inter_delay": 30},
    2: {"max_retries": 4, "wait_base": 15, "inter_delay": 10},
    3: {"max_retries": 3, "wait_base": 10, "inter_delay":  4},
}
_cfg = _TIER_CONFIG.get(API_TIER, _TIER_CONFIG[1])

MAX_RETRIES      = _cfg["max_retries"]
RETRY_WAIT_BASE  = _cfg["wait_base"]
INTER_FILE_DELAY = _cfg["inter_delay"]


# ─── BROKER DEFAULTS ────────────────────────────────────────────────

BROKER_DEFAULTS = {
    "name":  "Ιατρόπουλος Χρήστος",
    "tel":   "+30 697 590 0189",
    "email": "info@chiinsurancebrokers.com",
}
