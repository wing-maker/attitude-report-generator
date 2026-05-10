"""
image_downloader.py
Downloads images from Google Drive file IDs.
Uses the public download endpoint (no auth needed for "anyone with link" files).
"""

import os
import requests
from pathlib import Path
from drive_reader import extract_file_id


def download_drive_image(drive_url_or_id: str, output_path: str) -> bool:
    """
    Download a Google Drive image to output_path.
    Returns True on success, False on failure.

    Uses the public thumbnail/lh3 endpoint, which is more reliable than uc?export=download
    for files shared as "Anyone with the link".
    """
    if not drive_url_or_id:
        return False

    file_id = drive_url_or_id if "/" not in drive_url_or_id else extract_file_id(drive_url_or_id)
    if not file_id:
        return False

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    # Try multiple endpoints in order of reliability
    candidate_urls = [
        # 1. Thumbnail with very large size — redirects to public CDN
        f"https://drive.google.com/thumbnail?id={file_id}&sz=w2000",
        # 2. Direct CDN URL (sometimes works without redirect)
        f"https://lh3.googleusercontent.com/d/{file_id}=w2000",
        # 3. Old uc download endpoint as fallback
        f"https://drive.google.com/uc?export=download&id={file_id}",
    ]

    for url in candidate_urls:
        try:
            r = requests.get(url, timeout=30, stream=True, allow_redirects=True,
                             headers={"User-Agent": "Mozilla/5.0"})
            if r.status_code != 200:
                continue

            content_type = r.headers.get("Content-Type", "")
            # Accept any image MIME, reject HTML pages
            if "text/html" in content_type:
                # The "Google Drive — Virus scan warning" page for very large files
                import re
                html_text = r.text
                confirm_match = re.search(r'confirm=([^&"]+)', html_text)
                if confirm_match:
                    token = confirm_match.group(1)
                    url2 = f"https://drive.google.com/uc?export=download&confirm={token}&id={file_id}"
                    r = requests.get(url2, timeout=30, stream=True,
                                     headers={"User-Agent": "Mozilla/5.0"})
                    if r.status_code != 200:
                        continue
                else:
                    continue

            # Write file
            with open(output_path, "wb") as f:
                for chunk in r.iter_content(8192):
                    f.write(chunk)

            # Verify it's a real image (not zero-byte or HTML)
            if Path(output_path).stat().st_size < 1000:
                Path(output_path).unlink(missing_ok=True)
                continue

            return True

        except Exception as e:
            print(f"Download attempt error for {file_id} at {url[:80]}: {e}")
            continue

    return False


def download_all_screenshots(campaign_data: dict, work_dir: str) -> dict:
    """
    Downloads all post and insight screenshots for the entire campaign.
    Modifies campaign_data in place, adding local paths:
      - 'post_screenshot_local': path to downloaded post screenshot
      - 'insight_screenshot_locals': list of paths to downloaded insight screenshots
      - 'ig_post_screenshot_local': IG cross-post screenshot
      - 'ig_insight_screenshot_locals': list

    Returns dict with download stats: {'total': N, 'success': N, 'failed': N}
    """
    work_dir = Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)

    stats = {"total": 0, "success": 0, "failed": 0}

    def _safe_name(s, default):
        s = str(s or default)
        s = "".join(c if c.isalnum() or c in "_-" else "_" for c in s)
        return s[:40] or default

    def _download_one(url, out_path):
        stats["total"] += 1
        if download_drive_image(url, str(out_path)):
            stats["success"] += 1
            return str(out_path)
        else:
            stats["failed"] += 1
            return None

    for kind, kocs in [("xhs", campaign_data.get("xhs_kocs", [])),
                       ("tiktok", campaign_data.get("tiktok_kols", []))]:
        for i, koc in enumerate(kocs):
            base = f"{kind}_{i:02d}_{_safe_name(koc.get('creator'), 'creator')}"

            # Post screenshot
            koc["post_screenshot_local"] = None
            if koc.get("post_screenshot_url"):
                p = work_dir / f"{base}_post.png"
                koc["post_screenshot_local"] = _download_one(koc["post_screenshot_url"], p)

            # Insight screenshots (multiple)
            koc["insight_screenshot_locals"] = []
            for j, url in enumerate(koc.get("insight_screenshot_urls", [])):
                p = work_dir / f"{base}_insight_{j:02d}.png"
                local = _download_one(url, p)
                if local:
                    koc["insight_screenshot_locals"].append(local)

            # IG cross-post post screenshot
            koc["ig_post_screenshot_local"] = None
            if koc.get("ig_post_screenshot_url"):
                p = work_dir / f"{base}_ig_post.png"
                koc["ig_post_screenshot_local"] = _download_one(koc["ig_post_screenshot_url"], p)

            # IG cross-post insight screenshots
            koc["ig_insight_screenshot_locals"] = []
            for j, url in enumerate(koc.get("ig_insight_screenshot_urls", [])):
                p = work_dir / f"{base}_ig_insight_{j:02d}.png"
                local = _download_one(url, p)
                if local:
                    koc["ig_insight_screenshot_locals"].append(local)

    return stats
