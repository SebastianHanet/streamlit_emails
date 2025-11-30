"""
Streamlit application for sending bulk consulting outreach emails.

This variant lets you:

  * Upload a spreadsheet of leads (.xlsx or .csv)
  * Choose which column contains email addresses
  * Type a single subject line and email body
  * Either send all emails immediately OR schedule them to be sent later
  * Automatically attach your resume PDF to every email
  * Require a safety password before any emails are sent

Before running this app, make sure you have the following installed:

    pip install streamlit pandas openpyxl python-dotenv

For deployment (including Streamlit Community Cloud), set secrets in
``.env`` or ``.streamlit/secrets.toml`` so they stay on the server and
are never embedded in the client:

    SMTP_HOST=<your SMTP server>              # optional server-side default
    SMTP_PORT=<SMTP port, e.g. 587>           # optional server-side default
    SMTP_USER=<SMTP username>                 # optional server-side default
    SMTP_PASS=<SMTP password or app password> # optional server-side default
    USE_STARTTLS=True                         # optional server-side default
    OUTREACH_PASSWORD=<access code required before sending>  # required

Additionally, place your resume PDF (the file that will be attached to
every email) in the same folder as this script and set the
``RESUME_FILENAME`` constant below accordingly.

To run the app:

    streamlit run streamlit_app2.py
"""

import os
import threading
import time as time_module
from datetime import datetime
from email.message import EmailMessage
from typing import Tuple, List

import pandas as pd
import smtplib
import streamlit as st
from dotenv import load_dotenv

# Load environment variables from .env if present
load_dotenv()

# ---- Config ----
RESUME_FILENAME = "Sebastian Hanet Resume 2025.pdf"


def get_config_value(key: str, default: str = "") -> str:
    """Retrieve a secret or env var without exposing it to the UI."""
    try:
        return str(st.secrets[key])
    except Exception:
        return os.getenv(key, default)


def str_to_bool(value: str) -> bool:
    return str(value).strip().lower() in {"1", "true", "t", "yes", "y"}


# Safety password required before sending/scheduling emails (must be set in secrets)
OUTREACH_PASSWORD = get_config_value("OUTREACH_PASSWORD", "")


# ---- Utilities ----
def load_credentials_from_env() -> dict:
    """Load default SMTP creds/settings from environment."""
    port_raw = get_config_value("SMTP_PORT", "587")
    try:
        smtp_port = int(port_raw)
    except (TypeError, ValueError):
        smtp_port = 587

    return {
        "smtp_host": get_config_value("SMTP_HOST", ""),
        "smtp_port": smtp_port,
        "smtp_user": get_config_value("SMTP_USER", ""),
        "smtp_pass": get_config_value("SMTP_PASS", ""),
        # Do not auto-populate sender to avoid embedding addresses in the UI
        "sender_email": "",
        "use_starttls": str_to_bool(get_config_value("USE_STARTTLS", "True")),
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
    """Send an email with resume PDF attachment via SMTP."""
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


def send_bulk_emails(
    df: pd.DataFrame,
    email_column: str,
    subject: str,
    body: str,
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_pass: str,
    sender_email: str,
    use_starttls: bool,
) -> Tuple[int, int, List[str]]:
    """
    Send the same subject/body to every address in the given column.

    Returns (success_count, failure_count, error_messages).
    """
    success_count = 0
    failure_count = 0
    errors: List[str] = []

    for idx, row in df.iterrows():
        raw_email = row.get(email_column, "")
        to_email = str(raw_email).strip()
        if not to_email or "@" not in to_email:
            failure_count += 1
            errors.append(f"Row {idx + 1}: invalid email '{to_email}'")
            continue

        ok, msg = send_email(
            to_email=to_email,
            subject=subject,
            body=body,
            smtp_host=smtp_host,
            smtp_port=smtp_port,
            smtp_user=smtp_user,
            smtp_pass=smtp_pass,
            sender_email=sender_email,
            use_starttls=use_starttls,
        )

        if ok:
            success_count += 1
        else:
            failure_count += 1
            errors.append(f"Row {idx + 1} ({to_email}): {msg}")

        # Gentle throttle to avoid hammering the SMTP server
        time_module.sleep(0.2)

    return success_count, failure_count, errors


def schedule_bulk_email_send(
    send_at: datetime,
    df: pd.DataFrame,
    email_column: str,
    subject: str,
    body: str,
    smtp_host: str,
    smtp_port: int,
    smtp_user: str,
    smtp_pass: str,
    sender_email: str,
    use_starttls: bool,
) -> None:
    """
    Schedule a background job to send all emails at a future time.

    The Streamlit server process must stay running for this to execute.
    """
    delay_seconds = max((send_at - datetime.now()).total_seconds(), 0)

    def worker():
        time_module.sleep(delay_seconds)
        successes, failures, errors = send_bulk_emails(
            df=df,
            email_column=email_column,
            subject=subject,
            body=body,
            smtp_host=smtp_host,
            smtp_port=smtp_port,
            smtp_user=smtp_user,
            smtp_pass=smtp_pass,
            sender_email=sender_email,
            use_starttls=use_starttls,
        )
        # Basic logging to the server console
        print(
            f"[Bulk email job @ {send_at.isoformat()}] "
            f"Completed with {successes} successes and {failures} failures."
        )
        for err in errors:
            print("  -", err)

    threading.Thread(target=worker, daemon=True).start()


# ---- App ----
def main() -> None:
    st.title("Consulting Outreach Emailer - Bulk Sender")
    st.write(
        "Upload a spreadsheet of leads, choose the email column, "
        "type your outreach email once, and send it to everyone at once "
        "or schedule it for later."
    )

    if not OUTREACH_PASSWORD:
        st.error(
            "The app is locked until you set an OUTREACH_PASSWORD in "
            "Streamlit secrets or a .env file on the server."
        )
        return

    # Sidebar: SMTP configuration
    default_creds = load_credentials_from_env()
    st.sidebar.header("SMTP Settings")
    smtp_host_input = st.sidebar.text_input(
        "SMTP host", value="", placeholder="smtp.gmail.com", help="e.g. smtp.gmail.com"
    )
    smtp_port_input = st.sidebar.number_input(
        "SMTP port", value=default_creds["smtp_port"] or 587, step=1, format="%d"
    )
    use_starttls_input = st.sidebar.checkbox(
        "Use STARTTLS", value=default_creds["use_starttls"]
    )
    smtp_user_input = st.sidebar.text_input(
        "SMTP username", value="", placeholder="you@example.com"
    )
    smtp_pass_input = st.sidebar.text_input(
        "SMTP password / app password",
        value="",
        type="password",
    )
    sender_email_input = st.sidebar.text_input(
        "From address",
        value="",
        placeholder="you@example.com",
        help="Required each time to avoid embedding your address in the app.",
    )
    st.sidebar.caption(
        "Sensitive fields are never pre-filled. If SMTP_* secrets exist on the server, "
        "they are used when inputs are left blank (except the From address, which must "
        "always be provided manually)."
    )

    # Optional: quick SMTP sanity hint for Gmail
    with st.sidebar.expander("Gmail tip", expanded=False):
        st.markdown(
            "Use **smtp.gmail.com:587** with STARTTLS and a **Gmail App Password** "
            "(Google Account → Security → App passwords)."
        )

    # Resume preview in sidebar
    if os.path.exists(RESUME_FILENAME):
        with open(RESUME_FILENAME, "rb") as f:
            resume_bytes = f.read()
        st.sidebar.subheader("Attached Resume")
        st.sidebar.download_button(
            "Download attached resume",
            data=resume_bytes,
            file_name=os.path.basename(RESUME_FILENAME),
        )
    else:
        st.sidebar.error(
            f"Resume file '{RESUME_FILENAME}' not found in the app directory."
        )

    # File upload + persistence in session_state
    uploaded_file = st.file_uploader(
        "Upload spreadsheet (.xlsx or .csv)", type=["xlsx", "csv"], key="file_uploader"
    )

    df = None
    if uploaded_file is not None:
        try:
            if uploaded_file.name.lower().endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            st.session_state["df"] = df
            st.session_state["uploaded_filename"] = uploaded_file.name
        except Exception as exc:
            st.error(f"Could not read file: {exc}")
            return
    elif "df" in st.session_state:
        df = st.session_state["df"]

    if df is None:
        st.info(
            "Upload your spreadsheet to get started. At minimum you need one column "
            "containing email addresses."
        )
        return

    if df.empty:
        st.warning("Your spreadsheet is empty.")
        return

    st.subheader("Data preview")
    st.dataframe(df.head())

    # Email column selection
    email_cols = list(df.columns)
    default_email_col_index = 0
    for i, col in enumerate(email_cols):
        if "email" in str(col).lower():
            default_email_col_index = i
            break

    email_column = st.selectbox(
        "Column containing email addresses",
        options=email_cols,
        index=default_email_col_index,
    )

    # Single subject + body for all emails
    subject = st.text_input("Email subject")
    body = st.text_area(
        "Email body",
        height=320,
        help="This content will be sent to every recipient in the selected column.",
    )

    # Safety password input
    outreach_password = st.text_input(
        "Outreach password (required to send emails)",
        type="password",
        help="Enter the safety password before bulk sending.",
    )

    # Send now vs schedule
    send_mode = st.radio(
        "When should these emails be sent?",
        ["Send now", "Schedule for later"],
        horizontal=True,
    )

    scheduled_date = None
    scheduled_time = None
    if send_mode == "Schedule for later":
        col_date, col_time = st.columns(2)
        scheduled_date = col_date.date_input("Send date")
        scheduled_time = col_time.time_input("Send time")

        st.caption(
            "Emails will be sent using the server's local time. "
            "Keep this app running past the scheduled time."
        )

    if st.button("Create and send emails", type="primary"):
        sender_email = sender_email_input.strip()
        smtp_host = smtp_host_input.strip() or default_creds["smtp_host"]
        smtp_port = int(smtp_port_input or default_creds["smtp_port"] or 0)
        use_starttls = bool(use_starttls_input)
        smtp_user = smtp_user_input.strip() or default_creds["smtp_user"]
        smtp_pass = smtp_pass_input.strip() or default_creds["smtp_pass"]

        # Basic validation
        if not sender_email or "@" not in sender_email:
            st.error("Please set a valid 'From' address in the sidebar.")
            return

        if not smtp_host:
            st.error("Please provide an SMTP host (or set SMTP_HOST in secrets).")
            return

        if smtp_port <= 0:
            st.error("Please provide a valid SMTP port.")
            return

        if not smtp_user:
            st.error("Please provide an SMTP username or set SMTP_USER in secrets.")
            return

        if not smtp_pass:
            st.error("Please provide an SMTP password or set SMTP_PASS in secrets.")
            return

        if email_column not in df.columns:
            st.error("Selected email column is not present in the data.")
            return

        if not subject.strip():
            st.error("Please provide an email subject.")
            return

        if not body.strip():
            st.error("Please provide an email body.")
            return

        # Password check
        if outreach_password != OUTREACH_PASSWORD:
            st.error("Incorrect outreach password. Emails were not sent.")
            return

        if send_mode == "Send now":
            successes, failures, errors = send_bulk_emails(
                df=df,
                email_column=email_column,
                subject=subject.strip(),
                body=body,
                smtp_host=smtp_host.strip(),
                smtp_port=int(smtp_port),
                smtp_user=smtp_user.strip(),
                smtp_pass=smtp_pass.strip(),
                sender_email=sender_email.strip(),
                use_starttls=bool(use_starttls),
            )
            st.success(
                f"Bulk send finished. Success: {successes}, Failed: {failures}."
            )
            if errors:
                st.subheader("Errors")
                for err in errors:
                    st.write("-", err)
        else:
            # Schedule
            if scheduled_date is None or scheduled_time is None:
                st.error("Please choose a date and time for scheduled send.")
                return

            send_at = datetime.combine(scheduled_date, scheduled_time)
            if send_at <= datetime.now():
                st.error("Scheduled time must be in the future.")
                return

            schedule_bulk_email_send(
                send_at=send_at,
                df=df.copy(),
                email_column=email_column,
                subject=subject.strip(),
                body=body,
                smtp_host=smtp_host.strip(),
                smtp_port=int(smtp_port),
                smtp_user=smtp_user.strip(),
                smtp_pass=smtp_pass.strip(),
                sender_email=sender_email.strip(),
                use_starttls=bool(use_starttls),
            )
            st.info(
                f"Emails scheduled to be sent at {send_at} "
                "(server local time). Make sure this app stays running."
            )


if __name__ == "__main__":
    main()
