from flask import Flask, render_template, request, flash, redirect, url_for
import os
import pandas as pd
import yagmail
from config import EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_SUBJECT

app = Flask(__name__)
app.secret_key = "supersecretkey"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def send_certificate(name, email, certificate_path, message):
    try:
        yag = yagmail.SMTP(EMAIL_SENDER, EMAIL_PASSWORD)
        yag.send(
            to=email,
            subject=EMAIL_SUBJECT,
            contents=f"Hello {name},\n\n{message}\n\nBest Regards",
            attachments=certificate_path
        )
        yag.close()
        print(f"Email sent to {email} with {certificate_path}")
        return True
    except Exception as e:
        print(f"Error sending to {email}: {e}")
        return False

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        excel_file = request.files.get('excel_file')
        certificate_files = request.files.getlist('certificate_files')
        message = request.form.get('thank_you_message', '').strip()

        if not excel_file or not certificate_files:
            flash("Please upload both the Excel file and certificate PDFs.", "danger")
            return redirect(url_for('index'))

        # Save and read Excel file
        excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
        excel_file.save(excel_path)
        df = pd.read_excel(excel_path)

        # Normalize column names
        df.columns = df.columns.str.strip().str.lower()

        if "name" not in df.columns or "email" not in df.columns:
            flash("Excel file must have 'name' and 'email' columns.", "danger")
            return redirect(url_for('index'))

        # Save certificates and map them by filename
        certificate_paths = {}
        for file in certificate_files:
            filename = file.filename.strip().lower().replace(" ", "_")
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)
            certificate_paths[filename] = file_path

        print("Uploaded Certificates:", certificate_paths.keys())

        success_count, failed_count = 0, 0

        for _, row in df.iterrows():
            user_name = str(row["name"]).strip().lower().replace(" ", "_")
            user_email = str(row["email"]).strip()
            expected_certificate = f"{user_name}.pdf"

            print(f"Looking for {expected_certificate} in uploaded certificates.")

            if expected_certificate in certificate_paths:
                if send_certificate(row["name"], user_email, certificate_paths[expected_certificate], message):
                    success_count += 1
                else:
                    failed_count += 1
            else:
                print(f"Certificate not found for {user_name}")
                failed_count += 1

        flash(f"Certificates sent: {success_count}, Failed: {failed_count}", "info")
        return redirect(url_for('index'))

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
