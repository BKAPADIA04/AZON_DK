import io
import zipfile
import pandas as pd
import smtplib
from email.message import EmailMessage
import streamlit as st
import os

# ------------------------------
# 1. Preprocessing Function
# ------------------------------
def preprocess_flatno(flat_no: str) -> str:
    if not isinstance(flat_no, str):
        return ""
    flat_no = flat_no.upper().strip()
    flat_no = flat_no.replace("ROW HOUSE", "RH").replace("ROWHOUSE", "RH")
    flat_no = flat_no.replace(" ", "").replace("-", "")
    return flat_no

# ------------------------------
# 2. Load Excel Data
# ------------------------------
def load_excel(excel_file) -> pd.DataFrame:
    df = pd.read_excel(excel_file)
    df["FlatNo"] = df["FlatNo"].apply(preprocess_flatno)
    return df

# ------------------------------
# 3. Extract PDFs from ZIP
# ------------------------------
def extract_pdfs_from_zip(zip_file) -> list:
    pdf_files = []
    with zipfile.ZipFile(zip_file) as z:
        for file_name in z.namelist():
            if file_name.lower().endswith(".pdf"):
                pdf_data = z.read(file_name)
                pdf_file = io.BytesIO(pdf_data)
                pdf_file.name = os.path.basename(file_name)
                pdf_files.append(pdf_file)
    return pdf_files

# ------------------------------
# 4. Match PDFs with Flats
# ------------------------------
def collect_pdf_for_flat(flat_no: str, pdf_files: list) -> list:
    matched_files = []
    for pdf_file in pdf_files:
        normalized_name = preprocess_flatno(os.path.splitext(pdf_file.name)[0])
        if normalized_name == flat_no:
            matched_files.append(pdf_file)
    return matched_files

# ------------------------------
# 5. Send Email
# ------------------------------
def send_email(receiver_email, flat_no, attachments, subject, sender_email, password):
    body = f"""Dear Resident,

Please find attached your flat document for {flat_no}.

Regards,
Avon Plaza CHSL.
"""
    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg["Subject"] = subject
    msg.set_content(body)

    for file in attachments:
        file.seek(0)
        msg.add_attachment(file.read(), maintype="application", subtype="pdf", filename=file.name)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender_email, password)
        server.send_message(msg)

# ------------------------------
# 6. Streamlit App
# ------------------------------
def main():
    st.title("📧 Avon Plaza Bill Mailer")

    # Upload Excel
    excel_file = st.file_uploader("Upload Excel file with FlatNo & Email", type=["xlsx"])

    # Upload ZIP containing PDFs
    zip_file = st.file_uploader("Upload ZIP folder containing PDFs", type=["zip"])

    # Month + Year
    col1, col2 = st.columns(2)
    with col1:
        month = st.text_input("Enter month (e.g. AUG):").upper()
    with col2:
        year = st.text_input("Enter year (e.g. 25):")

    # Gmail settings (user input at runtime)
    sender_email = st.text_input("Sender Gmail address")
    password = st.text_input("Gmail App Password", type="password")

    if excel_file and zip_file and month and year and sender_email and password:
        if st.button("🚀 Send Emails"):
            subject = f"MAINTENANCE BILL & SUPPLEMENTARY BILL FOR THE MONTH OF {month} {year}"
            df = load_excel(excel_file)
            pdf_files = extract_pdfs_from_zip(zip_file)

            for _, row in df.iterrows():
                email = row["Email"]
                flat_no = row["FlatNo"]

                attachments = collect_pdf_for_flat(flat_no, pdf_files)

                if not attachments:
                    st.warning(f"[SKIP] No PDF found for {flat_no} ({email})")
                    continue

                send_email(email, flat_no, attachments, subject, sender_email, password)
                st.success(f"[SENT] Flat {flat_no} → {email}")

            st.balloons()
            st.info("✅ All mails processed.")

if __name__ == "__main__":
    main()
