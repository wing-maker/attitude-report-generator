"""
excel_editor.py
Exports campaign metrics to Excel for manual review.
"""

import pandas as pd


METRIC_COLUMNS = ["views", "likes", "comments", "saves", "shares"]


def campaign_rows(campaign_data: dict):
    rows = []

    for platform, list_name in (("XHS", "xhs_kocs"), ("TIKTOK", "tiktok_kols")):
        for idx, item in enumerate(campaign_data.get(list_name, [])):
            rows.append({
                "row_key": f"{list_name}:{idx}:main",
                "platform": platform,
                "post_type": "MAIN",
                "creator": item.get("creator", ""),
                "post_url": item.get("post_url", ""),
                "views": item.get("views", 0),
                "likes": item.get("likes", 0),
                "comments": item.get("comments", 0),
                "saves": item.get("saves", 0),
                "shares": item.get("shares", 0),
            })

            if item.get("ig_post_url"):
                rows.append({
                    "row_key": f"{list_name}:{idx}:ig",
                    "platform": "IG",
                    "post_type": f"{platform} EXTRA",
                    "creator": item.get("ig_creator") or item.get("creator", ""),
                    "post_url": item.get("ig_post_url", ""),
                    "views": item.get("ig_views", 0),
                    "likes": item.get("ig_likes", 0),
                    "comments": item.get("ig_comments", 0),
                    "saves": item.get("ig_saves", 0),
                    "shares": item.get("ig_shares", 0),
                })

    return rows


def export_metrics_excel(campaign_data: dict, output_path: str):
    df = pd.DataFrame(campaign_rows(campaign_data))

    summary = (
        df.groupby("platform", as_index=False)[METRIC_COLUMNS]
        .sum()
        .assign(posts=lambda x: df.groupby("platform").size().values)
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Metrics", index=False)
        summary.to_excel(writer, sheet_name="Summary Totals", index=False)

        for ws in writer.sheets.values():
            for col_cells in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col_cells)
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 60)

    return output_path
