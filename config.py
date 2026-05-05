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
}


def insurer_color(name: str) -> RGBColor:
    """Return a brand color for a known insurer, or teal as default."""
    n = name.upper()
    if "GENERALI" in n: return C["generali"]
    if "MORGAN"   in n: return C["navy"]
    if "NOW"      in n: return C["now"]
    if "ERGO"     in n: return rgb(0x00, 0x5A, 0xA0)
    if "AXA"      in n: return C["axa"]
    if "ALLIANZ"  in n: return C["allianz"]
    if "CIGNA"    in n: return C["cigna"]
    return C["teal"]


# ─── API / RATE-LIMIT SETTINGS ──────────────────────────────────────

MODEL            = "claude-sonnet-4-20250514"
MAX_RETRIES      = 3
RETRY_WAIT_BASE  = 10   # seconds; doubles each retry: 10 → 20 → 40
INTER_FILE_DELAY = 4    # seconds to pause between consecutive PDF calls


# ─── BROKER DEFAULTS ────────────────────────────────────────────────

BROKER_DEFAULTS = {
    "name":  "Ιατρόπουλος Χρήστος",
    "tel":   "+30 697 590 0189",
    "email": "info@chiinsurancebrokers.com",
}
