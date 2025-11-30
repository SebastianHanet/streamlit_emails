"""
Streamlit application for sending personalized consulting emails.

This app allows you to upload a spreadsheet of consulting leads and their
pre-written email templates, review and edit each email (including the
recipient address and subject line), and send the messages one at a time
with your resume attached. After you send or skip a message, the app
automatically advances to the next row until all emails have been
processed.

Before running this app, make sure you have the following installed:

    pip install streamlit pandas openpyxl python-dotenv

You should also create a ``.env`` file in the same directory with the
following environment variables defined (these are used as defaults in
the UI but can be changed in the sidebar):

    SENDER_EMAIL=<your email address>
    SMTP_HOST=<your SMTP server>
    SMTP_PORT=<SMTP port, e.g. 587>
    SMTP_USER=<SMTP username>
    SMTP_PASS=<SMTP password or app password>
    USE_STARTTLS=True

Additionally, place your resume PDF (the file that will be attached to
every email) in the same folder as this script and set the
``RESUME_FILENAME`` constant below accordingly.

To run the app:

    streamlit run streamlit_app.py

The app does not automatically send any emails without your explicit
action. Each time you click “Send Email” the message will be sent via
SMTP using the credentials provided.
"""

import os
from email.message import EmailMessage
from typing import Tuple

import pandas as pd
import smtplib
import streamlit as st
from dotenv import load_dotenv

# Load environment variables from .env if present
load_dotenv()

# ---- Config ----
RESUME_FILENAME = "Sebastian Hanet Resume 2025.pdf"


# ---- Utilities ----
def _rerun():
    """Streamlit rerun with backward/forward compatibility."""
    if hasattr(st, "rerun"):
        st.rerun()
    elif hasattr(st, "experimental_rerun"):
        st.experimental_rerun()


def parse_template(template: str) -> Tuple[str, str]:
    """
    Split a template into subject/body.

    Looks for a first line starting with "Subject:" and returns (subject, body).
    If none found, returns ("", full_template).
    """
    template = template or ""
    lines = template.splitlines()
    subject = ""
    body_start = 0
    if lines and lines[0].lower().startswith("subject:"):
        subject = lines[0].split(":", 1)[1].strip()
        body_start = 1
        # skip any blank lines after subject
        while body_start < len(lines) and not lines[body_start].strip():
            body_start += 1
    body = "\n".join(lines[body_start:])
    return subject, body


def load_credentials_from_env() -> dict:
    """Load default SMTP creds/settings from environment."""
    return {
        "smtp_host": os.getenv("SMTP_HOST", ""),
        "smtp_port": int(os.getenv("SMTP_PORT", "587")),
        "smtp_user": os.getenv("SMTP_USER", ""),
        "smtp_pass": os.getenv("SMTP_PASS", ""),
        "sender_email": os.getenv("SENDER_EMAIL", ""),
        "use_starttls": os.getenv("USE_STARTTLS", "True").lower() == "true",
    }


def send_email(
    to_email: str,
    subject: str,
    body: str,
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_pass: str,
    sender_email: str,
    use_starttls: bool,
) -> Tuple[bool, str]:
    """Send an email with optional PDF attachment via SMTP."""
    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)

    # Attach resume if present
    if os.path.exists(RESUME_FILENAME):
        with open(RESUME_FILENAME, "rb") as f:
            resume_data = f.read()
        msg.add_attachment(
            resume_data,
            maintype="application",
            subtype="pdf",
            filename=os.path.basename(RESUME_FILENAME),
        )
    else:
        return False, f"Resume file '{RESUME_FILENAME}' not found in app folder."

    try:
        if use_starttls:
            with smtplib.SMTP(smtp_host, smtp_port) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(smtp_user, smtp_pass)
                server.send_message(msg)
        else:
            with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
                server.login(smtp_user, smtp_pass)
                server.send_message(msg)
        return True, f"Email sent to {to_email}."
    except Exception as exc:
        return False, str(exc)


# ---- Row/State helpers ----
def _ensure_df_loaded(uploaded_file):
    """
    Keep the uploaded file and DataFrame stable across reruns.
    Only load into session_state when a *new* file is uploaded.
    """
    if uploaded_file is None:
        return

    if (
        "df" not in st.session_state
        or st.session_state.get("uploaded_filename") != uploaded_file.name
    ):
        df = pd.read_excel(uploaded_file)
        st.session_state["df"] = df
        st.session_state["uploaded_filename"] = uploaded_file.name
        st.session_state["row_idx"] = 0
        # reset field state so first row preloads correctly
        for k in ("to", "subject", "body", "last_loaded_idx"):
            st.session_state.pop(k, None)


def _derive_subject_for_row(row: pd.Series) -> str:
    """
    Prefer an explicit 'Subject' column if present/non-empty.
    Else parse from template, else fall back to a safe default.
    """
    # If the DF has a Subject column and it's set, use it
    if "Subject" in row and isinstance(row["Subject"], str) and row["Subject"].strip():
        return row["Subject"].strip()

    # Else try to parse from template
    tpl = str(row.get("Email Template", "") or "")
    subj, _ = parse_template(tpl)
    if subj.strip():
        return subj.strip()

    # Fallback
    return "Intro from Sebastian - independent data science consultant"


def preload_fields(idx: int):
    """Load current row values into session_state exactly once per row."""
    df = st.session_state["df"]
    row = df.iloc[idx]

    st.session_state["to"] = str(row.get("Email", "") or "")
    st.session_state["subject"] = _derive_subject_for_row(row)
    st.session_state["body"] = str(row.get("Email Template", "") or "")


def advance_to_next_row():
    """Increment row index, clear widget state, rerun."""
    st.session_state["row_idx"] += 1
    for k in ("to", "subject", "body"):
        st.session_state.pop(k, None)
    _rerun()


# ---- App ----
def main() -> None:
    st.title("Consulting Email Sender")
    st.write(
        "Upload a spreadsheet of consulting leads with their email templates, "
        "review and edit each message, and send them one by one."
    )

    # Sidebar: SMTP configuration
    default_creds = load_credentials_from_env()
    st.sidebar.header("SMTP Settings")
    smtp_host = st.sidebar.text_input(
        "SMTP host", value=default_creds["smtp_host"], help="e.g. smtp.gmail.com"
    )
    smtp_port = st.sidebar.number_input("SMTP port", value=default_creds["smtp_port"], step=1, format="%d")
    use_starttls = st.sidebar.checkbox("Use STARTTLS", value=default_creds["use_starttls"])
    smtp_user = st.sidebar.text_input("SMTP username", value=default_creds["smtp_user"])
    smtp_pass = st.sidebar.text_input("SMTP password / app password", value=default_creds["smtp_pass"], type="password")
    sender_email = st.sidebar.text_input("From address", value=default_creds["sender_email"])

    # Optional: quick SMTP sanity hint for Gmail
    with st.sidebar.expander("Gmail tip", expanded=False):
        st.markdown(
            "Use **smtp.gmail.com:587** with STARTTLS and a **Gmail App Password** "
            "(Google Account → Security → App passwords)."
        )

    # Upload
    uploaded_file = st.file_uploader("Upload spreadsheet (.xlsx)", type=["xlsx"], key="file_uploader")
    _ensure_df_loaded(uploaded_file)

    if "df" not in st.session_state:
        st.info("Upload your spreadsheet to get started. Expected columns: Company, Email, Email Template (and optional Subject).")
        # Resume preview
        if os.path.exists(RESUME_FILENAME):
            with open(RESUME_FILENAME, "rb") as f:
                resume_bytes = f.read()
            st.sidebar.subheader("Attached Resume")
            st.sidebar.download_button("Download attached resume", data=resume_bytes, file_name=os.path.basename(RESUME_FILENAME))
        else:
            st.sidebar.error(f"Resume file '{RESUME_FILENAME}' not found in the app directory.")
        return

    df = st.session_state["df"]
    n = len(df)
    if n == 0:
        st.warning("Your spreadsheet is empty.")
        return

    # Current row index
    if "row_idx" not in st.session_state:
        st.session_state["row_idx"] = 0

    idx = st.session_state["row_idx"]
    if idx >= n:
        st.success("All emails have been processed. 🎉")
        # Resume preview still available
        if os.path.exists(RESUME_FILENAME):
            with open(RESUME_FILENAME, "rb") as f:
                resume_bytes = f.read()
            st.sidebar.subheader("Attached Resume")
            st.sidebar.download_button("Download attached resume", data=resume_bytes, file_name=os.path.basename(RESUME_FILENAME))
        else:
            st.sidebar.error(f"Resume file '{RESUME_FILENAME}' not found in the app directory.")
        return

    # Preload the widgets when we land on a new row
    if st.session_state.get("last_loaded_idx") != idx:
        preload_fields(idx)
        st.session_state["last_loaded_idx"] = idx

    row = df.iloc[idx]
    st.caption(f"Row {idx + 1} of {n}")
    st.subheader(f"{row.get('Company', '')}")

    # Editable fields bound to session_state so they repopulate correctly
    to_val = st.text_input("Recipient email", key="to")
    subj_val = st.text_input("Subject", key="subject")
    body_val = st.text_area("Email body", key="body", height=320)

    # Actions
    col1, col2 = st.columns(2)
    send_clicked = col1.button("Send Email", type="primary")
    skip_clicked = col2.button("Skip")

    if send_clicked:
        # Basic validation
        if not to_val or "@" not in to_val:
            st.error("Please provide a valid recipient email.")
        elif not sender_email or "@" not in sender_email:
            st.error("Please set a valid 'From' address in the sidebar.")
        else:
            ok, msg = send_email(
                to_email=to_val.strip(),
                subject=subj_val.strip() or _derive_subject_for_row(row),
                body=body_val,
                smtp_host=smtp_host.strip(),
                smtp_port=int(smtp_port),
                smtp_user=smtp_user.strip(),
                smtp_pass=smtp_pass.strip(),
                sender_email=sender_email.strip(),
                use_starttls=bool(use_starttls),
            )
            if ok:
                st.toast("Email sent ✅", icon="✅")
                advance_to_next_row()
            else:
                st.error(f"Failed to send email: {msg}")

    if skip_clicked:
        st.info("Skipped.")
        advance_to_next_row()

    # Progress + resume preview
    st.progress((idx + 1) / max(1, n))
    if os.path.exists(RESUME_FILENAME):
        with open(RESUME_FILENAME, "rb") as f:
            resume_bytes = f.read()
        st.sidebar.subheader("Attached Resume")
        st.sidebar.download_button("Download attached resume", data=resume_bytes, file_name=os.path.basename(RESUME_FILENAME))
    else:
        st.sidebar.error(f"Resume file '{RESUME_FILENAME}' not found in the app directory.")


if __name__ == "__main__":
    main()
