"""
excel_editor.py
Exports campaign metrics to Excel and applies corrected metrics back to data.
"""

import pandas as pd


METRIC_COLUMNS = ["views", "likes", "comments", "saves", "shares"]


def _to_int(value):
    if pd.isna(value):
        return 0
    if isinstance(value, str):
        value = value.replace(",", "").strip()
        if not value:
            return 0
    return int(float(value))


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
    rows = campaign_rows(campaign_data)
    df = pd.DataFrame(rows)

    summary = (
        df.groupby("platform", as_index=False)[METRIC_COLUMNS]
        .sum()
        .assign(posts=lambda x: df.groupby("platform").size().values)
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Edit Metrics", index=False)
        summary.to_excel(writer, sheet_name="Summary Totals", index=False)

        ws = writer.sheets["Edit Metrics"]
        for col_cells in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 60)

        ws_summary = writer.sheets["Summary Totals"]
        for col_cells in ws_summary.columns:
            max_len = max(len(str(cell.value or "")) for cell in col_cells)
            ws_summary.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 24)

    return output_path


def apply_metrics_excel(campaign_data: dict, excel_file):
    df = pd.read_excel(excel_file, sheet_name="Edit Metrics")
    required = {"row_key", *METRIC_COLUMNS}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(sorted(missing))}")

    for _, row in df.iterrows():
        row_key = str(row["row_key"])
        try:
            list_name, idx_text, mode = row_key.split(":")
            item = campaign_data[list_name][int(idx_text)]
        except Exception:
            continue

        if mode == "ig":
            prefix = "ig_"
        else:
            prefix = ""

        for metric in METRIC_COLUMNS:
            item[f"{prefix}{metric}"] = _to_int(row[metric])

    return campaign_data
