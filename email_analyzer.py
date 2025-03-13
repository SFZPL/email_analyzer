import os
import time
from datetime import datetime, timedelta, time as dt_time
import streamlit as st
import tempfile
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import openai
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()  # Loads variables from .env into the environment

# Gmail integration scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# Configure OpenAI (standard API)
openai.api_key = os.getenv("OPENAI_API_KEY")  # Ensure this is set in your .env file

def get_gmail_service():
    if "gmail_creds" not in st.session_state:
        # Load client config from Streamlit secrets
        client_config = st.secrets["gcp"]["client_config"]
        # Write the JSON string to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as temp:
            temp.write(client_config.encode("utf-8"))
            temp_path = temp.name
        # Run the OAuth flow for the user
        flow = InstalledAppFlow.from_client_secrets_file(temp_path, SCOPES)
        creds = flow.run_local_server(port=0)
        st.session_state.gmail_creds = creds
    else:
        creds = st.session_state.gmail_creds
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            st.session_state.gmail_creds = creds
    service = build('gmail', 'v1', credentials=creds)
    return service

def fetch_recent_emails(service, start_datetime, end_datetime, desired_count=50, max_fetch=200):
    """
    Fetch emails by scanning up to max_fetch messages and returning those whose internalDate
    falls between start_datetime and end_datetime.
    """
    matching_emails = []
    next_page_token = None
    fetched_messages = 0

    st.write("Fetching emails...")
    while fetched_messages < max_fetch and len(matching_emails) < desired_count:
        params = {"userId": "me", "maxResults": 50}
        if next_page_token:
            params["pageToken"] = next_page_token
        result = service.users().messages().list(**params).execute()
        messages = result.get("messages", [])
        fetched_messages += len(messages)
        if not messages:
            break
        for msg in messages:
            msg_data = service.users().messages().get(userId="me", id=msg["id"]).execute()
            # internalDate is in milliseconds since epoch
            internal_date = int(msg_data.get("internalDate", 0))
            email_datetime = datetime.fromtimestamp(internal_date / 1000)
            # Check if email is within the desired range
            if start_datetime <= email_datetime <= end_datetime:
                snippet = msg_data.get("snippet", "")
                headers = msg_data["payload"].get("headers", [])
                subject = next((h["value"] for h in headers if h["name"] == "Subject"), "No Subject")
                sender = next((h["value"] for h in headers if h["name"] == "From"), "Unknown Sender")
                matching_emails.append({
                    "id": msg["id"],
                    "subject": subject,
                    "from": sender,
                    "snippet": snippet,
                    "received_at": email_datetime,
                })
                if len(matching_emails) >= desired_count:
                    break
        next_page_token = result.get("nextPageToken")
        if not next_page_token:
            break
    st.write(f"Fetched {len(matching_emails)} emails in the selected range out of {fetched_messages} messages scanned.")
    return matching_emails

def analyze_email_openai(email_text: str) -> str:
    """
    Uses OpenAI's ChatCompletion with gpt-4-turbo to extract structured info from an email.
    """
    prompt = (
        f"Extract all relevant details for a service request from the email below:\n\n"
        f"{email_text}\n\n"
        "Return the extracted information in a clear, structured format."
    )
    
    max_retries = 5
    wait_time = 10  # seconds
    for attempt in range(max_retries):
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4-turbo",
                messages=[
                    {"role": "system", "content": "You are an expert in extracting structured information from client emails."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=200,
                temperature=0.3,
                top_p=0.95,
                frequency_penalty=0,
                presence_penalty=0,
            )
            return response.choices[0].message.content.strip()
        except openai.error.RateLimitError:
            st.write("Rate limit exceeded, retrying...")
            if attempt < max_retries - 1:
                time.sleep(wait_time)
            else:
                raise

def compute_priority(email, extracted_info):
    # Calculate urgency score by counting urgency-related keywords in the extracted info.
    urgency_keywords = ['urgent', 'asap', 'immediately', 'critical']
    urgency_score = sum(1 for word in urgency_keywords if word in extracted_info.lower())
    
    # If the email is from PWC, give it a high bonus to ensure it is prioritized.
    bonus = 100 if "pwc" in email["from"].lower() else 0
    
    total_score = bonus + urgency_score
    return total_score

def main():
    st.title("Overnight Email Summary Dashboard")
    
    # Default date/time: yesterday at 5:00 PM to today at 9:00 AM
    today = datetime.today().date()
    yesterday = today - timedelta(days=1)
    default_start_time = dt_time(17, 0)  # 5:00 PM
    default_end_time = dt_time(9, 0)     # 9:00 AM

    st.subheader("Select Date and Time Range for Emails")
    start_date = st.date_input("Start Date", value=yesterday)
    start_time = st.time_input("Start Time", value=default_start_time)
    end_date = st.date_input("End Date", value=today)
    end_time = st.time_input("End Time", value=default_end_time)

    # Combine date and time into datetime objects (local time)
    start_datetime = datetime.combine(start_date, start_time)
    end_datetime = datetime.combine(end_date, end_time)

    st.write("Analyzing emails from:", start_datetime, "to", end_datetime)

    if st.button("Run Analysis"):
        service = get_gmail_service()
        emails = fetch_recent_emails(service, start_datetime, end_datetime, desired_count=50, max_fetch=200)

        if not emails:
            st.write("No emails found in the selected date/time range.")
            return

        processed_emails = []
        progress_bar = st.progress(0)
        st.write("Starting email analysis...")
        for i, email in enumerate(emails):
            st.write(f"Processing email {i+1} of {len(emails)}: {email['subject']}")
            analysis = analyze_email_openai(email["snippet"])
            priority = compute_priority(email, analysis)
            email.update({
                "analysis": analysis,
                "priority": priority,
            })
            processed_emails.append(email)
            progress_bar.progress((i+1)/len(emails))
        
        sorted_emails = sorted(processed_emails, key=lambda x: x["priority"], reverse=True)
        
        st.write("### Prioritized Email Summaries")
        for email in sorted_emails:
            st.write(f"**Subject:** {email['subject']}")
            st.write(f"**From:** {email['from']}")
            st.write(f"**Received At:** {email['received_at']}")
            st.write(f"**Priority Score:** {email['priority']}")
            st.write(f"**Summary:** {email['analysis']}")
            st.markdown("---")

if __name__ == '__main__':
    main()
