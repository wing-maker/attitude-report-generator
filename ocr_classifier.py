"""
ocr_classifier.py
Identifies the "correct" insight screenshot from a list of multiple uploads.
Uses Tesseract OCR (free, local) + keyword matching against known platform UI patterns.
"""

from PIL import Image

try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False


# =============================================================
# Keyword sets for each platform
# =============================================================
# XHS "correct" page = "笔记数据" / "Note Analysis" overview screen
XHS_STRONG_KEYWORDS = ["笔记数据", "Note Analysis", "Overview"]
XHS_METRIC_KEYWORDS = [
    "曝光数", "观看数", "点赞数", "评论数", "收藏数", "分享数",
    "Likes", "Views", "Comments", "Saves", "Shares", "Impressions"
]

# TikTok "correct" page = "Video analysis" + "Key metrics" Overview tab
TIKTOK_STRONG_KEYWORDS = ["Key metrics", "Video views", "TikTok Studio", "Video analysis", "Total play time"]
TIKTOK_METRIC_KEYWORDS = ["Likes", "Views", "Comments", "Shares", "Saves"]


def ocr_text(image_path: str, lang: str = "chi_sim+eng") -> str:
    """Run OCR on an image. Returns extracted text (empty string if OCR unavailable)."""
    if not TESSERACT_AVAILABLE:
        return ""
    try:
        img = Image.open(image_path)
        return pytesseract.image_to_string(img, lang=lang)
    except Exception as e:
        print(f"OCR error on {image_path}: {e}")
        return ""


def score_image(image_path: str, platform: str) -> int:
    """
    Score how likely this image is the "correct" insight overview.
    Higher = more likely correct.

    Scoring:
      Strong keyword match: +3 points each
      Metric keyword match: +1 point each
    """
    if platform == "xhs":
        text = ocr_text(image_path, lang="chi_sim+eng")
        strong = XHS_STRONG_KEYWORDS
        metrics = XHS_METRIC_KEYWORDS
    elif platform == "tiktok":
        text = ocr_text(image_path, lang="eng")
        strong = TIKTOK_STRONG_KEYWORDS
        metrics = TIKTOK_METRIC_KEYWORDS
    else:
        return 0

    score = 0
    for kw in strong:
        if kw.lower() in text.lower():
            score += 3
    for kw in metrics:
        if kw.lower() in text.lower():
            score += 1

    return score


def find_correct_insight(image_paths: list, platform: str) -> tuple:
    """
    Returns: (best_path, confidence_level, all_scores)
      confidence_level: 'high' | 'low' | 'none'
      all_scores: list of (path, score) for UI display

    'high'  = clear winner — auto-use it, don't bother user
    'low'   = ambiguous — ask user to pick
    'none'  = OCR not available — default to first image
    """
    if not image_paths:
        return None, "none", []

    if not TESSERACT_AVAILABLE:
        return image_paths[0], "none", [(p, 0) for p in image_paths]

    if len(image_paths) == 1:
        # Only one option — use it but mark confidence by checking it has keywords
        score = score_image(image_paths[0], platform)
        confidence = "high" if score >= 3 else "low"
        return image_paths[0], confidence, [(image_paths[0], score)]

    # Score all images
    scored = [(p, score_image(p, platform)) for p in image_paths]
    scored_sorted = sorted(scored, key=lambda x: -x[1])

    best_path, best_score = scored_sorted[0]
    second_score = scored_sorted[1][1] if len(scored_sorted) > 1 else 0

    # Decide confidence
    # High confidence: clear winner with score >= 6 AND beats second by >= 2
    if best_score >= 6 and (best_score - second_score) >= 2:
        confidence = "high"
    elif best_score >= 4:
        confidence = "low"  # has some signal but not clear-cut
    else:
        confidence = "low"  # nothing clearly matches

    return best_path, confidence, scored


def classify_all_insights(campaign_data: dict) -> list:
    """
    Iterates through all KOCs/KOLs and identifies the correct insight screenshot.
    Modifies campaign_data in place:
      - sets 'best_insight_local' to the chosen path
      - sets 'insight_confidence' to 'high'/'low'/'none'

    Returns: list of low-confidence cases that need user review:
      [{'group_key': ..., 'creator': ..., 'platform': ..., 'options': [(path, score), ...], 'current_pick': path}, ...]
    """
    low_confidence = []

    for kind, kocs in [("xhs", campaign_data.get("xhs_kocs", [])),
                       ("tiktok", campaign_data.get("tiktok_kols", []))]:
        for i, koc in enumerate(kocs):
            paths = koc.get("insight_screenshot_locals", [])
            best, conf, scores = find_correct_insight(paths, kind)
            koc["best_insight_local"] = best
            koc["insight_confidence"] = conf
            koc["insight_scores"] = scores

            if conf == "low" and len(paths) > 1:
                low_confidence.append({
                    "group_key": f"{kind}_{i}",
                    "creator": koc.get("creator", "?"),
                    "platform": kind,
                    "options": scores,
                    "current_pick": best,
                })

            # IG cross-post insights — also classify (use IG's platform for scoring? Use 'xhs'-like for both since IG insights look similar)
            ig_paths = koc.get("ig_insight_screenshot_locals", [])
            ig_best, ig_conf, ig_scores = find_correct_insight(ig_paths, "xhs")  # IG insights tend to use similar UI
            koc["best_ig_insight_local"] = ig_best
            koc["ig_insight_confidence"] = ig_conf
            koc["ig_insight_scores"] = ig_scores

    return low_confidence
