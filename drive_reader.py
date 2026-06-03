"""
drive_reader.py
Reads a public Google Drive folder and finds XHS / TIKTOK subfolders + their response sheets.
No service account needed — uses public CSV export endpoint.
"""

import re
import requests
import pandas as pd
from io import StringIO


# =============================================================
# Helpers to extract Drive IDs
# =============================================================
def extract_folder_id(url: str) -> str:
    """Extract folder ID from a Drive URL like:
    https://drive.google.com/drive/folders/XXX?usp=...
    """
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
    if not m:
        raise ValueError("Invalid Drive folder URL — must be like https://drive.google.com/drive/folders/...")
    return m.group(1)


def extract_file_id(url: str) -> str:
    """Extract file ID from any Drive URL (open?id=, /file/d/, /spreadsheets/d/)."""
    if not url or not isinstance(url, str):
        return ""
    url = url.strip()
    patterns = [
        r"id=([a-zA-Z0-9_-]+)",
        r"/file/d/([a-zA-Z0-9_-]+)",
        r"/spreadsheets/d/([a-zA-Z0-9_-]+)",
        r"/folders/([a-zA-Z0-9_-]+)",
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return ""


# =============================================================
# Public folder listing (uses Drive's no-auth endpoint)
# =============================================================
def list_folder_contents(folder_id: str):
    """
    Lists files/folders in a public Drive folder using the unauthenticated viewer endpoint.
    Returns list of dicts: [{'id': ..., 'name': ..., 'is_folder': bool, 'mime_type': ...}]

    Uses the public 'embeddedfolderview' approach. If folder is not public, raises ValueError.
    """
    # Drive's embedded folder list endpoint (works without auth IF folder is "Anyone with link")
    url = f"https://drive.google.com/embeddedfolderview?id={folder_id}#list"
    r = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
    if r.status_code != 200:
        raise ValueError(
            f"Cannot access folder. Make sure the folder is shared as "
            f"'Anyone with the link can view'. (HTTP {r.status_code})"
        )

    # Force UTF-8 to handle Chinese folder/file names correctly
    r.encoding = "utf-8"
    html = r.text
    # The HTML contains rows like:
    # <a href="https://drive.google.com/folderview?id=XXX" ... class="flip-entry-list-item">
    #   <div class="flip-entry-info">
    #     <div class="flip-entry-title">XHS</div>
    items = []

    # The actual HTML structure looks like:
    # <a href="...folders/XXX" target="_blank">
    #   ...nested divs...
    #   <div class="flip-entry-title">NAME</div>
    # </a>
    # So we use BeautifulSoup-style: find each <a>, then look for flip-entry-title inside.

    # Folders: extract id + title separately
    folder_anchor_pattern = re.compile(
        r'<a[^>]+href="([^"]*?/drive/folders/([a-zA-Z0-9_-]+)[^"]*)"[^>]*>(.*?)</a>',
        re.DOTALL
    )
    file_anchor_pattern = re.compile(
        r'<a[^>]+href="([^"]*?/file/d/([a-zA-Z0-9_-]+)[^"]*)"[^>]*>(.*?)</a>',
        re.DOTALL
    )
    sheet_anchor_pattern = re.compile(
        r'<a[^>]+href="([^"]*?/spreadsheets/d/([a-zA-Z0-9_-]+)[^"]*)"[^>]*>(.*?)</a>',
        re.DOTALL
    )

    title_pattern = re.compile(
        r'<div class="flip-entry-title"[^>]*>([^<]+)</div>',
        re.DOTALL
    )

    def extract_title(inner_html: str) -> str:
        """Extract the flip-entry-title text from inside an <a> tag's inner HTML."""
        m = title_pattern.search(inner_html)
        if m:
            return m.group(1).strip()
        # Fallback: strip HTML tags and take any leftover text
        clean = re.sub(r'<[^>]+>', '', inner_html).strip()
        return clean

    for m in folder_anchor_pattern.finditer(html):
        href = m.group(1)
        folder_id = m.group(2)
        inner = m.group(3)
        title = extract_title(inner)
        if title:
            items.append({"id": folder_id, "name": title, "is_folder": True, "type": "folder", "href": href})

    for m in file_anchor_pattern.finditer(html):
        href = m.group(1)
        file_id = m.group(2)
        inner = m.group(3)
        title = extract_title(inner)
        if title:
            items.append({"id": file_id, "name": title, "is_folder": False, "type": "file", "href": href})

    for m in sheet_anchor_pattern.finditer(html):
        href = m.group(1)
        sheet_id = m.group(2)
        inner = m.group(3)
        title = extract_title(inner)
        if title:
            items.append({"id": sheet_id, "name": title, "is_folder": False, "type": "sheet", "href": href})

    # Deduplicate (same id might appear multiple times if there are nested anchors)
    seen = set()
    unique = []
    for it in items:
        if it["id"] not in seen:
            seen.add(it["id"])
            unique.append(it)

    return unique


def resolve_shortcut_folder_id(file_id: str) -> str | None:
    """
    Best-effort resolver for public Google Drive folder shortcuts.
    Embedded folder view can list shortcuts as files, while the public shortcut
    page often redirects to or contains the target /drive/folders/<id> URL.
    """
    urls = [
        f"https://drive.google.com/file/d/{file_id}/view",
        f"https://drive.google.com/open?id={file_id}",
    ]
    headers = {"User-Agent": "Mozilla/5.0"}

    for url in urls:
        try:
            r = requests.get(url, timeout=15, headers=headers, allow_redirects=True)
        except requests.RequestException:
            continue

        final_match = re.search(r"/drive/folders/([a-zA-Z0-9_-]+)", r.url)
        if final_match:
            return final_match.group(1)

        body_match = re.search(r"/drive/folders/([a-zA-Z0-9_-]+)", r.text)
        if body_match:
            return body_match.group(1)

    return None


def resolve_platform_folder_item(item: dict) -> tuple[str | None, str | None]:
    """Return (folder_id, warning) for a folder or folder shortcut listing item."""
    if item["is_folder"]:
        return item["id"], None

    folder_id = resolve_shortcut_folder_id(item["id"])
    if folder_id:
        return folder_id, None

    return None, f"Could not open response shortcut: {item['name']}"


# =============================================================
# Find platform subfolders (XHS / TIKTOK)
# =============================================================
def find_platform_subfolders(root_folder_id: str):
    """
    Returns dict: {'xhs': folder_id_or_None, 'tiktok': folder_id_or_None}
    Matching is case-insensitive on folder name.
    """
    items = list_folder_contents(root_folder_id)
    result = {"xhs": None, "tiktok": None, "warnings": []}

    for it in items:
        name_lower = it["name"].lower()
        if "xhs" not in name_lower and "rednote" not in name_lower and "tiktok" not in name_lower and "tt" != name_lower:
            continue

        folder_id, warning = resolve_platform_folder_item(it)
        if warning:
            result["warnings"].append(warning)
            continue
        if "xhs" in name_lower or "rednote" in name_lower or "小红书" in it["name"]:
            result["xhs"] = folder_id
        if "tiktok" in name_lower or "tt" == name_lower:
            result["tiktok"] = folder_id

    return result


# =============================================================
# Find response sheet inside a platform folder
# =============================================================
def find_response_sheet_id(folder_id: str):
    """
    Inside a platform folder, finds the Google Sheet (form responses).
    Returns the spreadsheet ID, or None if not found.
    """
    items = list_folder_contents(folder_id)
    for it in items:
        if it["is_folder"]:
            continue
        # Sheet name typically contains "Responses" but could be just a Sheet
        # The embedded folder view doesn't reveal mime types, so we just take the first non-folder
        # Heuristic: name contains "Response" or "Sheet" or just take the first non-folder file
        if "response" in it["name"].lower() or "sheet" in it["name"].lower():
            return it["id"]

    # Fallback: if there's exactly 1 non-folder, use it
    files = [it for it in items if not it["is_folder"]]
    if len(files) >= 1:
        return files[0]["id"]

    return None


# =============================================================
# Read response sheet as CSV
# =============================================================
def read_sheet_as_dataframe(sheet_id: str) -> pd.DataFrame:
    """Reads a public Google Sheet (Anyone with link) as a pandas DataFrame using CSV export."""
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    r = requests.get(url, timeout=20)
    if r.status_code != 200:
        raise ValueError(
            f"Cannot read response sheet. Ensure it's shared as 'Anyone with link can view'. (HTTP {r.status_code})"
        )
    # Force UTF-8 encoding to handle Chinese characters correctly
    r.encoding = "utf-8"
    df = pd.read_csv(StringIO(r.text))
    return df


# =============================================================
# Parse a response row into a structured dict
# =============================================================
def parse_xhs_row(row: pd.Series) -> dict:
    """
    Parse one row of the XHS response sheet.
    Expected columns (based on Jombeli form structure):
      0  Timestamp
      1  Email Address
      2  Creator Name (XHS)
      3  XHS Posted Link
      4  XHS Post Screenshot (drive link)
      5  XHS Insight Screenshot (drive link, may have multiple URLs comma/space separated)
      6  Views
      7  Likes
      8  Comments
      9  Save
      10 Share
      11 Impressions
      12 IG Profile Name
      13 IG Posted Link
      14 IG Post Screenshot
      15 IG Insight Screenshot
      16 Views
      17 Likes
      18 Comments
      19 Saves
      20 Shares
    """
    cols = list(row.values)

    def to_int(v):
        try:
            if pd.isna(v): return 0
            return int(float(str(v).replace(",", "").strip()))
        except (ValueError, TypeError):
            return 0

    def split_links(s):
        """Insight screenshot field may contain multiple URLs. Split them."""
        if pd.isna(s) or not s: return []
        # URLs are typically separated by comma or whitespace
        parts = re.split(r'[,\s]+', str(s).strip())
        return [p for p in parts if p.startswith("http")]

    return {
        "platform": "xhs",
        "creator": str(cols[2]) if not pd.isna(cols[2]) else "",
        "post_url": str(cols[3]) if not pd.isna(cols[3]) else "",
        "post_screenshot_url": str(cols[4]) if not pd.isna(cols[4]) else "",
        "insight_screenshot_urls": split_links(cols[5]),
        "views": to_int(cols[6]),
        "likes": to_int(cols[7]),
        "comments": to_int(cols[8]),
        "saves": to_int(cols[9]),
        "shares": to_int(cols[10]),
        # Cross-post on IG
        "ig_creator": str(cols[12]) if len(cols) > 12 and not pd.isna(cols[12]) else "",
        "ig_post_url": str(cols[13]) if len(cols) > 13 and not pd.isna(cols[13]) else "",
        "ig_post_screenshot_url": str(cols[14]) if len(cols) > 14 and not pd.isna(cols[14]) else "",
        "ig_insight_screenshot_urls": split_links(cols[15]) if len(cols) > 15 else [],
        "ig_views": to_int(cols[16]) if len(cols) > 16 else 0,
        "ig_likes": to_int(cols[17]) if len(cols) > 17 else 0,
        "ig_comments": to_int(cols[18]) if len(cols) > 18 else 0,
        "ig_saves": to_int(cols[19]) if len(cols) > 19 else 0,
        "ig_shares": to_int(cols[20]) if len(cols) > 20 else 0,
    }


def parse_tiktok_row(row: pd.Series) -> dict:
    """Same column structure as XHS — see parse_xhs_row docs."""
    d = parse_xhs_row(row)
    d["platform"] = "tiktok"
    return d


# =============================================================
# High-level: fetch entire campaign data
# =============================================================
def fetch_campaign_data(folder_url: str):
    """
    Main entry point.
    Returns dict:
      {
        'xhs_kocs': [...],      # list of parsed dicts
        'tiktok_kols': [...],
        'errors': [...],        # list of warning messages
      }
    """
    folder_id = extract_folder_id(folder_url)
    subfolders = find_platform_subfolders(folder_id)

    result = {"xhs_kocs": [], "tiktok_kols": [], "errors": []}
    result["errors"].extend(subfolders.get("warnings", []))

    # XHS
    if subfolders["xhs"]:
        sheet_id = find_response_sheet_id(subfolders["xhs"])
        if sheet_id:
            try:
                df = read_sheet_as_dataframe(sheet_id)
                for _, row in df.iterrows():
                    parsed = parse_xhs_row(row)
                    if parsed["creator"]:
                        result["xhs_kocs"].append(parsed)
            except Exception as e:
                result["errors"].append(f"XHS sheet read error: {e}")
        else:
            result["errors"].append("XHS response sheet not found")
    else:
        result["errors"].append("XHS subfolder not found")

    # TikTok
    if subfolders["tiktok"]:
        sheet_id = find_response_sheet_id(subfolders["tiktok"])
        if sheet_id:
            try:
                df = read_sheet_as_dataframe(sheet_id)
                for _, row in df.iterrows():
                    parsed = parse_tiktok_row(row)
                    if parsed["creator"]:
                        result["tiktok_kols"].append(parsed)
            except Exception as e:
                result["errors"].append(f"TikTok sheet read error: {e}")
        else:
            result["errors"].append("TikTok response sheet not found")
    else:
        result["errors"].append("TikTok subfolder not found")

    return result


# =============================================================
# Quick test
# =============================================================
if __name__ == "__main__":
    test_url = "https://drive.google.com/drive/folders/1eYlzXk-qfnLLaI3ayTecHpCB-cm-r4HJ"
    data = fetch_campaign_data(test_url)
    print(f"XHS KOCs: {len(data['xhs_kocs'])}")
    print(f"TikTok KOLs: {len(data['tiktok_kols'])}")
    print(f"Errors: {data['errors']}")
    if data['xhs_kocs']:
        print("\nFirst XHS:")
        for k, v in data['xhs_kocs'][0].items():
            print(f"  {k}: {v}")
