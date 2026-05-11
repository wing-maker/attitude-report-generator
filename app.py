"""
app.py
Main Streamlit application — Attitude Campaign Report Generator
Run with: streamlit run app.py
"""

import streamlit as st
import os
import tempfile
from pathlib import Path
from PIL import Image

import drive_reader
import image_downloader
import ocr_classifier
import color_extractor
import pptx_builder
import drive_uploader
import excel_editor
import apps_script_uploader


# =============================================================
# Page config
# =============================================================
st.set_page_config(
    page_title="Attitude Campaign Report Generator",
    page_icon="📊",
    layout="centered",
)

# =============================================================
# Constants
# =============================================================
TEMPLATE_PATH = "template.pptx"
BRANDS_DIR = Path("brands")
DEFAULT_HEADER_COLOR = (0, 160, 168)  # Watsons teal — most campaigns are Watsons


# =============================================================
# Helpers
# =============================================================
def list_brand_logos():
    """List all .png/.jpg files in the brands/ folder."""
    if not BRANDS_DIR.exists():
        return []
    files = sorted([
        f for f in BRANDS_DIR.iterdir()
        if f.suffix.lower() in (".png", ".jpg", ".jpeg")
    ])
    return [f.stem.title() for f in files], [f for f in files]


def get_google_oauth_config():
    """Read Google OAuth settings from Streamlit secrets."""
    secret_paths = [
        Path.home() / ".streamlit" / "secrets.toml",
        Path.cwd() / ".streamlit" / "secrets.toml",
    ]
    if not any(path.exists() for path in secret_paths):
        return None

    try:
        config = st.secrets.get("google_oauth", {})
    except Exception:
        return None

    required = ("client_id", "client_secret", "redirect_uri")
    if not all(config.get(k) for k in required):
        return None
    return {
        "client_id": config["client_id"],
        "client_secret": config["client_secret"],
        "redirect_uri": config["redirect_uri"],
    }


def get_google_service_account_config():
    """Read Google service account JSON from Streamlit secrets."""
    try:
        config = st.secrets.get("google_service_account", {})
    except Exception:
        return None

    required = ("client_email", "private_key", "token_uri")
    if not all(config.get(k) for k in required):
        return None
    return dict(config)


def get_apps_script_upload_config():
    """Read Apps Script upload endpoint settings from Streamlit secrets."""
    try:
        config = st.secrets.get("apps_script_upload", {})
    except Exception:
        return None

    required = ("web_app_url", "token")
    if not all(config.get(k) for k in required):
        return None
    return {
        "web_app_url": config["web_app_url"],
        "token": config["token"],
    }


def get_query_param(name: str):
    value = st.query_params.get(name)
    if isinstance(value, list):
        return value[0] if value else None
    return value


def handle_google_drive_callback():
    """Handle Google OAuth redirect no matter which app step is showing."""
    code = get_query_param("code")
    if not code or "drive_credentials" in st.session_state:
        return

    oauth_config = get_google_oauth_config()
    if not oauth_config:
        return

    try:
        st.session_state.drive_credentials = drive_uploader.exchange_code_for_credentials(
            client_id=oauth_config["client_id"],
            client_secret=oauth_config["client_secret"],
            redirect_uri=oauth_config["redirect_uri"],
            code=code,
        )
        st.session_state.drive_connected_message = "Google Drive connected."
        st.query_params.clear()
    except Exception as e:
        st.session_state.drive_connection_error = f"Could not connect Google Drive: {e}"


def show_drive_upload_section(output_path: str, upload_name: str, metrics_path: str, metrics_name: str):
    st.markdown("**Save to Google Drive**")
    st.caption("Upload the generated PPTX and metrics Excel back into the same campaign folder.")

    apps_script_config = get_apps_script_upload_config()
    if apps_script_config:
        if st.button("Upload PPTX + Metrics Excel to Same Drive Folder", use_container_width=True):
            try:
                folder_id = drive_reader.extract_folder_id(st.session_state.folder_url)
                uploaded_ppt = apps_script_uploader.upload_file(
                    web_app_url=apps_script_config["web_app_url"],
                    token=apps_script_config["token"],
                    folder_id=folder_id,
                    file_path=output_path,
                    file_name=upload_name,
                    mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
                uploaded_excel = apps_script_uploader.upload_file(
                    web_app_url=apps_script_config["web_app_url"],
                    token=apps_script_config["token"],
                    folder_id=folder_id,
                    file_path=metrics_path,
                    file_name=metrics_name,
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success("Uploaded PPTX and metrics Excel to Google Drive.")
                st.markdown(f"[Open uploaded report]({uploaded_ppt['webViewLink']})")
                st.markdown(f"[Open uploaded metrics Excel]({uploaded_excel['webViewLink']})")
            except Exception as e:
                st.error(f"Upload failed: {e}")
        return

    service_account_config = get_google_service_account_config()
    if service_account_config:
        if st.button("Upload PPTX + Metrics Excel to Same Drive Folder", use_container_width=True):
            try:
                folder_id = drive_reader.extract_folder_id(st.session_state.folder_url)
                uploaded_ppt = drive_uploader.upload_file_to_folder(
                    credentials_info=service_account_config,
                    folder_id=folder_id,
                    file_path=output_path,
                    file_name=upload_name,
                    mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_service_account=True,
                )
                uploaded_excel = drive_uploader.upload_file_to_folder(
                    credentials_info=service_account_config,
                    folder_id=folder_id,
                    file_path=metrics_path,
                    file_name=metrics_name,
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_service_account=True,
                )
                st.success("Uploaded PPTX and metrics Excel to Google Drive.")
                st.markdown(f"[Open uploaded report]({uploaded_ppt['webViewLink']})")
                st.markdown(f"[Open uploaded metrics Excel]({uploaded_excel['webViewLink']})")
            except Exception as e:
                st.error(f"Upload failed: {e}")
        return

    oauth_config = get_google_oauth_config()
    if not oauth_config:
        st.info(
            "Google Drive upload is not configured yet. Add "
            "`google_oauth.client_id`, `google_oauth.client_secret`, and "
            "`google_oauth.redirect_uri` in Streamlit secrets."
        )
        return

    if "drive_credentials" not in st.session_state:
        try:
            auth_url = drive_uploader.make_auth_url(**oauth_config)
            st.link_button("Connect Google Drive", auth_url, use_container_width=True)
            st.caption("After Google login, you will return here and can upload the report.")
        except Exception as e:
            st.error(f"Could not start Google Drive login: {e}")
        return

    if st.button("Upload PPTX + Metrics Excel to Same Drive Folder", use_container_width=True):
        try:
            folder_id = drive_reader.extract_folder_id(st.session_state.folder_url)
            uploaded_ppt = drive_uploader.upload_file_to_folder(
                credentials_info=st.session_state.drive_credentials,
                folder_id=folder_id,
                file_path=output_path,
                file_name=upload_name,
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            uploaded_excel = drive_uploader.upload_file_to_folder(
                credentials_info=st.session_state.drive_credentials,
                folder_id=folder_id,
                file_path=metrics_path,
                file_name=metrics_name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success("Uploaded PPTX and metrics Excel to Google Drive.")
            st.markdown(f"[Open uploaded report]({uploaded_ppt['webViewLink']})")
            st.markdown(f"[Open uploaded metrics Excel]({uploaded_excel['webViewLink']})")
        except Exception as e:
            st.error(f"Upload failed: {e}")


def show_metrics_excel_download(download_name: str):
    st.markdown("**Review Metrics**")
    st.caption("Download an Excel copy of the metrics for checking.")

    excel_path = os.path.join(st.session_state.work_dir, "editable_metrics.xlsx")
    excel_editor.export_metrics_excel(st.session_state.campaign_data, excel_path)

    metrics_name = download_name.replace(".pptx", "_Metrics.xlsx")
    with open(excel_path, "rb") as f:
        st.download_button(
            "Download Editable Metrics Excel",
            data=f.read(),
            file_name=metrics_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    return excel_path, metrics_name


# =============================================================
# Session state init
# =============================================================
handle_google_drive_callback()

if "step" not in st.session_state:
    st.session_state.step = "input"  # input -> processing -> review -> done
if "campaign_data" not in st.session_state:
    st.session_state.campaign_data = None
if "low_confidence" not in st.session_state:
    st.session_state.low_confidence = []
if "header_color" not in st.session_state:
    st.session_state.header_color = DEFAULT_HEADER_COLOR
if "work_dir" not in st.session_state:
    st.session_state.work_dir = None


# =============================================================
# Sidebar — Attitude branding
# =============================================================
with st.sidebar:
    st.markdown("### **ATTITUDE**")
    st.caption("IDEOLOGY")
    st.divider()
    st.markdown("**Campaign Report Generator**")
    st.caption("Internal tool for generating posting reports from Google Drive form responses.")
    st.divider()
    if st.button("🔄 Start Over", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()


# =============================================================
# Header
# =============================================================
st.title("📊 Campaign Report Generator")
st.caption("Generate a posting report from a Google Drive folder in 30 seconds.")

if st.session_state.get("drive_connected_message"):
    st.success(st.session_state.pop("drive_connected_message"))
if st.session_state.get("drive_connection_error"):
    st.error(st.session_state.pop("drive_connection_error"))


# =============================================================
# STEP 1 — INPUT FORM
# =============================================================
if st.session_state.step == "input":
    with st.form("input_form"):
        st.subheader("Step 1: Campaign Info")

        folder_url = st.text_input(
            "📁 Campaign Drive Folder Link",
            placeholder="https://drive.google.com/drive/folders/...",
            help="The folder must be shared as 'Anyone with link can view'."
        )

        col1, col2 = st.columns(2)
        with col1:
            campaign_name = st.text_input("📝 Campaign Name", placeholder="Jombeli Ampang Point")
        with col2:
            campaign_month = st.text_input("📅 Campaign Month", placeholder="Apr 2026")

        # Client logo selection
        brand_names, brand_paths = list_brand_logos()
        client_choice = st.selectbox(
            "🏪 Client Logo",
            options=brand_names + ["📁 Upload custom logo..."] if brand_names else ["📁 Upload custom logo..."],
            help="Choose a saved logo or upload your own."
        )

        custom_logo_file = None
        if client_choice == "📁 Upload custom logo...":
            custom_logo_file = st.file_uploader("Upload PNG/JPG logo", type=["png", "jpg", "jpeg"])

        # Hero image (optional)
        st.markdown("**🎨 Campaign Hero Image** (optional — used for cover and to derive header color)")
        hero_image_file = st.file_uploader("Upload hero image", type=["png", "jpg", "jpeg"], key="hero_uploader")

        submitted = st.form_submit_button("🚀 Generate Report", type="primary", use_container_width=True)

    if submitted:
        # Validate
        errors = []
        if not folder_url.strip():
            errors.append("Please enter a Drive folder link")
        if not campaign_name.strip():
            errors.append("Please enter a campaign name")
        if not campaign_month.strip():
            errors.append("Please enter a campaign month")
        if client_choice == "📁 Upload custom logo..." and not custom_logo_file:
            errors.append("Please upload a custom logo or select an existing client")

        if errors:
            for e in errors:
                st.error(e)
        else:
            # Save inputs to session
            st.session_state.folder_url = folder_url.strip()
            st.session_state.campaign_name = campaign_name.strip()
            st.session_state.campaign_month = campaign_month.strip()

            # Resolve client logo path
            if client_choice == "📁 Upload custom logo...":
                # Save uploaded file
                tmp_dir = tempfile.mkdtemp(prefix="arg_")
                logo_path = Path(tmp_dir) / "client_logo.png"
                with open(logo_path, "wb") as f:
                    f.write(custom_logo_file.getbuffer())
                st.session_state.client_logo_path = str(logo_path)
            else:
                idx = brand_names.index(client_choice)
                st.session_state.client_logo_path = str(brand_paths[idx])

            # Save hero image
            st.session_state.hero_image_path = None
            if hero_image_file:
                tmp_dir = tempfile.mkdtemp(prefix="arg_hero_")
                hero_path = Path(tmp_dir) / "hero.png"
                with open(hero_path, "wb") as f:
                    f.write(hero_image_file.getbuffer())
                st.session_state.hero_image_path = str(hero_path)

            st.session_state.step = "processing"
            st.rerun()


# =============================================================
# STEP 2 — PROCESSING
# =============================================================
if st.session_state.step == "processing":
    st.subheader("Step 2: Processing...")

    progress_bar = st.progress(0)
    status = st.empty()

    try:
        # 1. Read Drive folder
        status.info("📂 Reading campaign Drive folder...")
        progress_bar.progress(10)
        data = drive_reader.fetch_campaign_data(st.session_state.folder_url)

        if data["errors"]:
            for e in data["errors"]:
                st.warning(f"⚠️ {e}")

        if not data["xhs_kocs"] and not data["tiktok_kols"]:
            st.error("❌ No campaign data found. Check that the folder is shared correctly.")
            if st.button("← Back"):
                st.session_state.step = "input"
                st.rerun()
            st.stop()

        st.success(f"✅ Found {len(data['xhs_kocs'])} XHS + {len(data['tiktok_kols'])} TikTok submissions")

        # 2. Download all screenshots
        status.info("📷 Downloading screenshots... this may take a minute")
        progress_bar.progress(30)
        work_dir = tempfile.mkdtemp(prefix="arg_work_")
        st.session_state.work_dir = work_dir

        stats = image_downloader.download_all_screenshots(data, os.path.join(work_dir, "screenshots"))
        st.success(f"✅ Downloaded {stats['success']}/{stats['total']} screenshots")
        if stats['failed'] > 0:
            st.warning(f"⚠️ {stats['failed']} screenshots failed to download — slides will show placeholders.")

        # 3. Run OCR to identify correct insight images
        status.info("🔍 Identifying correct insight screenshots (OCR)...")
        progress_bar.progress(60)
        low_conf = ocr_classifier.classify_all_insights(data)
        st.session_state.low_confidence = low_conf

        # 4. Extract color from hero image (if provided)
        if st.session_state.hero_image_path:
            status.info("🎨 Extracting brand color from hero image...")
            progress_bar.progress(80)
            st.session_state.header_color = color_extractor.get_header_color_from_hero(
                st.session_state.hero_image_path
            )

        st.session_state.campaign_data = data
        progress_bar.progress(100)
        status.success("✅ Processing complete!")

        # If there are low-confidence insight cases, go to review step
        if low_conf:
            st.session_state.step = "review"
        else:
            st.session_state.step = "build"
        st.rerun()

    except Exception as e:
        st.error(f"❌ Error: {e}")
        if st.button("← Back"):
            st.session_state.step = "input"
            st.rerun()


# =============================================================
# STEP 3 — REVIEW (only if there are low-confidence cases)
# =============================================================
if st.session_state.step == "review":
    st.subheader("Step 3: Confirm Insight Screenshots")
    st.info(f"⚠️ Found {len(st.session_state.low_confidence)} cases where the correct insight image isn't clear. Please confirm:")

    # Color picker (if hero was provided)
    if st.session_state.hero_image_path:
        st.markdown("**🎨 Header bar color** (extracted from hero image):")
        col1, col2 = st.columns([1, 4])
        with col1:
            color_hex = color_extractor.rgb_to_hex(st.session_state.header_color)
            st.color_picker("Color", color_hex, key="color_picker_input", label_visibility="collapsed")
        with col2:
            st.caption("Adjust if needed, then continue below.")

        # Update color from picker
        new_hex = st.session_state.color_picker_input
        st.session_state.header_color = color_extractor.hex_to_rgb(new_hex)

    st.divider()

    # Iterate through low-confidence cases
    for case in st.session_state.low_confidence:
        st.markdown(f"#### {case['creator']} ({case['platform'].upper()})")
        st.caption("Click the correct image (the one showing complete data: views, likes, comments, saves, shares)")

        options = case["options"]  # list of (path, score)

        # Show as columns of small images
        cols = st.columns(len(options))
        for i, (path, score) in enumerate(options):
            with cols[i]:
                # Create a small thumbnail for display (full image too big for Streamlit)
                try:
                    from PIL import Image as PILImage
                    import io
                    img = PILImage.open(path)
                    img.thumbnail((400, 600))  # max 400x600 px for preview
                    buf = io.BytesIO()
                    # Convert to RGB if it's RGBA (some PNGs have alpha that breaks thumbnails)
                    if img.mode in ("RGBA", "LA", "P"):
                        img = img.convert("RGB")
                    img.save(buf, format="JPEG", quality=70)
                    buf.seek(0)
                    st.image(buf.getvalue())
                except Exception as e:
                    st.caption(f"(preview failed: {type(e).__name__}: {str(e)[:80]})")
                st.caption(f"Score: {score}")

                if st.radio(
                    f"Pick #{i+1}",
                    ["Pick this"],
                    key=f"radio_{case['group_key']}_{i}",
                    label_visibility="collapsed",
                    index=None,
                ):
                    pass  # handled below via session

        # Use a selectbox to make the choice (simpler than custom radio)
        labels = [f"Image {i+1} (score: {score})" for i, (_, score) in enumerate(options)]
        default_idx = next((i for i, (p, _) in enumerate(options) if p == case["current_pick"]), 0)
        choice = st.selectbox(
            f"Selected for {case['creator']}:",
            options=labels,
            index=default_idx,
            key=f"select_{case['group_key']}",
        )
        chosen_idx = labels.index(choice)
        chosen_path = options[chosen_idx][0]

        # Update the data
        platform = case["platform"]
        idx_in_list = int(case["group_key"].split("_")[1])
        if platform == "xhs":
            st.session_state.campaign_data["xhs_kocs"][idx_in_list]["best_insight_local"] = chosen_path
        else:
            st.session_state.campaign_data["tiktok_kols"][idx_in_list]["best_insight_local"] = chosen_path

        st.divider()

    if st.button("✅ Confirm & Build Report", type="primary", use_container_width=True):
        st.session_state.step = "build"
        st.rerun()


# =============================================================
# STEP 4 — BUILD
# =============================================================
if st.session_state.step == "build":
    st.subheader("Building Report...")

    progress = st.progress(0)
    status = st.empty()

    try:
        # Update color from picker if user adjusted
        if st.session_state.hero_image_path and "color_picker_input" in st.session_state:
            st.session_state.header_color = color_extractor.hex_to_rgb(
                st.session_state.color_picker_input
            )

        status.info("📝 Building PowerPoint...")
        progress.progress(50)

        output_path = os.path.join(st.session_state.work_dir, "report.pptx")
        pptx_builder.build_report(
            template_path=TEMPLATE_PATH,
            output_path=output_path,
            campaign_name=st.session_state.campaign_name,
            campaign_month=st.session_state.campaign_month,
            client_logo_path=st.session_state.client_logo_path,
            hero_image_path=st.session_state.hero_image_path,
            header_color=st.session_state.header_color,
            campaign_data=st.session_state.campaign_data,
        )

        progress.progress(100)
        status.success("✅ Report ready!")

        # Generate download filename
        safe_campaign = "".join(c if c.isalnum() else "_" for c in st.session_state.campaign_name)
        safe_month = "".join(c if c.isalnum() else "_" for c in st.session_state.campaign_month)
        download_name = f"{safe_campaign}_Posting_Report_{safe_month}.pptx"

        with open(output_path, "rb") as f:
            st.download_button(
                "📥 Download Report",
                data=f.read(),
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
            )

        metrics_path, metrics_name = show_metrics_excel_download(download_name)

        show_drive_upload_section(output_path, download_name, metrics_path, metrics_name)

        st.divider()

        # Show summary stats
        data = st.session_state.campaign_data
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Posts", len(data["xhs_kocs"]) + len(data["tiktok_kols"]))
        with col2:
            total_views = sum(k.get("views", 0) for k in data["xhs_kocs"] + data["tiktok_kols"])
            st.metric("Total Views (main)", f"{total_views:,}")
        with col3:
            total_eng = sum(
                k.get("likes", 0) + k.get("comments", 0) + k.get("saves", 0) + k.get("shares", 0)
                for k in data["xhs_kocs"] + data["tiktok_kols"]
            )
            st.metric("Total Engagement (main)", f"{total_eng:,}")

        if st.button("🔄 Generate Another Report"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

    except Exception as e:
        st.error(f"❌ Error building report: {e}")
        import traceback
        st.code(traceback.format_exc())
        if st.button("← Back to start"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
