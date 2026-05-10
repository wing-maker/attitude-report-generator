"""
color_extractor.py
Extracts the dominant color from a campaign hero image.
Adjusts brightness/contrast for use as a header bar (white text must be readable).
"""

from PIL import Image


def extract_dominant_color(image_path: str) -> tuple:
    """
    Returns the dominant RGB color as a tuple (r, g, b).
    Uses PIL's quantize for simplicity (no extra dependency).
    """
    try:
        img = Image.open(image_path).convert("RGB")
        # Resize for speed
        img.thumbnail((150, 150))
        # Get top color via quantize (palette of 5 colors)
        result = img.quantize(colors=5).convert("RGB")
        # Get color with most pixels
        colors = result.getcolors(maxcolors=10000)
        if not colors:
            return (0, 160, 168)  # default Watsons teal
        colors.sort(key=lambda x: -x[0])

        # Skip near-white and near-black (they're often borders)
        for count, (r, g, b) in colors:
            brightness = (r + g + b) / 3
            saturation = (max(r, g, b) - min(r, g, b))
            if 30 < brightness < 230 and saturation > 30:
                return (r, g, b)

        # Fallback: return first
        return colors[0][1]
    except Exception as e:
        print(f"Color extraction error: {e}")
        return (0, 160, 168)


def ensure_readable_for_white_text(rgb: tuple) -> tuple:
    """
    Adjusts color so white text is readable on it.
    If color is too light, darkens it. Returns adjusted RGB tuple.
    """
    r, g, b = rgb
    # Calculate perceived brightness (W3C formula)
    brightness = (r * 299 + g * 587 + b * 114) / 1000

    # If too bright, darken proportionally
    if brightness > 160:
        factor = 160 / brightness
        r = int(r * factor)
        g = int(g * factor)
        b = int(b * factor)

    # Clamp
    return (max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))


def rgb_to_hex(rgb: tuple) -> str:
    """Convert (r, g, b) tuple to hex string '#RRGGBB'."""
    return "#{:02X}{:02X}{:02X}".format(*rgb)


def hex_to_rgb(hex_str: str) -> tuple:
    """Convert '#RRGGBB' or 'RRGGBB' to (r, g, b)."""
    s = hex_str.lstrip("#")
    return tuple(int(s[i:i+2], 16) for i in (0, 2, 4))


def get_header_color_from_hero(hero_image_path: str) -> tuple:
    """
    Main entry point.
    Given a campaign hero image, returns the recommended header bar RGB color.
    Returns Watsons teal (0, 160, 168) if hero is None or extraction fails.
    """
    if not hero_image_path:
        return (0, 160, 168)

    raw = extract_dominant_color(hero_image_path)
    return ensure_readable_for_white_text(raw)
