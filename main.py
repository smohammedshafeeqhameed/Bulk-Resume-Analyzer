
import os
import google.generativeai as genai
import pandas as pd
import re
from PyPDF2 import PdfReader
from docx import Document
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configure the Gemini API key
genai.configure(api_key='API_KEY')

def get_gemini_feedback(text):
    """Analyzes resume text using the Gemini API and returns feedback."""
    model = genai.GenerativeModel('models/gemini-pro-latest')
    prompt = f""" You are an expert career coach. Review the following resume text and give concise, actionable feedback:
    - Summarize key strengths (1–2 points)
    - List specific improvement suggestions (3–5 bullet points)
    - Mention how the candidate can improve formatting or clarity.
    Resume text: {text[:8000]}  # limit to avoid token overload
    """
    response = model.generate_content(prompt)
    return response.text

def get_text_from_pdf(file_path):
    """Extracts text from a PDF file."""
    with open(file_path, 'rb') as f:
        reader = PdfReader(f)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    return text

def get_text_from_docx(file_path):
    """Extracts text from a DOCX file."""
    doc = Document(file_path)
    text = ''
    for para in doc.paragraphs:
        text += para.text + '\n'
    return text

def find_email(text):
    """Finds the first email address in the given text."""
    match = re.search(r'[\w.+-]+@[\w-]+\.[\w.-]+', text)
    return match.group(0) if match else None

def send_email(to_email, subject, body):
    """Sends an email using SMTP."""
    # Replace with your email and password
    from_email = 'email_id'
    password = 'passcode'

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, password)
        server.send_message(msg)
        server.quit()
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")


def main():
    """Main function to process resumes and send feedback."""
    resume_folder = 'resumes'
    feedback_data = []

    for filename in os.listdir(resume_folder):
        file_path = os.path.join(resume_folder, filename)
        text = ''
        if filename.endswith('.pdf'):
            text = get_text_from_pdf(file_path)
        elif filename.endswith('.docx'):
            text = get_text_from_docx(file_path)
        else:
            continue

        email = find_email(text)
        if not email:
            print(f"Could not find email in {filename}")
            continue

        print(f"Processing {filename} for {email}...")
        feedback = get_gemini_feedback(text)
        feedback_data.append([filename, email, feedback])

        # Send email with feedback
        subject = 'Resume Feedback'
        body = f"""Hello,

Here is some feedback on your resume:

{feedback}

Best regards,
The Resume Review Team"""
        send_email(email, subject, body)


    # Create a Pandas DataFrame and save it to an Excel file
    df = pd.DataFrame(feedback_data, columns=['Filename', 'Email', 'Feedback'])
    df.to_excel('resume_feedback.xlsx', index=False)
    print("Resume feedback saved to resume_feedback.xlsx")

if __name__ == '__main__':
    main()
