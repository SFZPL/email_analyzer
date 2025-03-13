import os
import time
import json
from datetime import datetime, timedelta, time as dt_time
import streamlit as st
import tempfile
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import openai
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()  # Loads variables from .env into the environment

# Gmail integration scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

openai.api_key = os.getenv("OPENAI_API_KEY", "")
if not openai.api_key and "openai" in st.secrets:
    openai.api_key = st.secrets["openai"]["api_key"]

def get_gmail_service():
    """
    Handles Gmail authentication flow through Google OAuth.
    Uses Streamlit session state to maintain credentials between reruns.
    """
    # Debug tab to help troubleshoot authentication issues
    with st.expander("Authentication Debugging (Expand if having issues)"):
        st.write("Session state keys:", list(st.session_state.keys()))
        query_params = st.query_params
        st.write("Query parameters:", query_params)
    
    # Check if we already have credentials
    if "gmail_creds" in st.session_state:
        creds = st.session_state.gmail_creds
        # Refresh token if expired
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                st.session_state.gmail_creds = creds
            except Exception as e:
                st.error(f"Error refreshing credentials: {e}")
                # Clear credentials to restart auth flow
                del st.session_state.gmail_creds
                st.rerun()
    else:
        # Load client config from Streamlit secrets
        try:
            client_config_str = st.secrets["gcp"]["client_config"]
            client_config = json.loads(client_config_str)
            
            # Write the client config to a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as temp:
                temp.write(json.dumps(client_config).encode("utf-8"))
                temp_path = temp.name
            
            # Use the hardcoded redirect URI that matches your Google Cloud Console configuration
            redirect_uri = "https://emailanalyzer-jeepcuohhmah2mqp8x3gqb.streamlit.app/"
            
            st.write(f"Using redirect URI: {redirect_uri}")
            
            # Create the OAuth flow
            flow = InstalledAppFlow.from_client_secrets_file(
                temp_path,
                SCOPES,
                redirect_uri=redirect_uri
            )
            
            # Check for authorization code in query parameters
            query_params = st.query_params
            if "code" in query_params:
                try:
                    # Get the authorization code
                    code = query_params["code"]
                    st.write("Attempting to exchange code for token...")
                    
                    # Exchange code for tokens
                    flow.fetch_token(code=code)
                    st.session_state.gmail_creds = flow.credentials
                    
                    # Clean up the URL by removing the query parameters
                    # Note: This might not work in all Streamlit environments
                    try:
                        st.set_query_params()
                    except:
                        pass
                        
                    st.success("Authentication successful!")
                    time.sleep(1)  # Give a moment for the success message to display
                    st.rerun()  # Rerun to clear the auth parameters from URL
                except Exception as e:
                    st.error(f"Error exchanging code for token: {str(e)}")
                    st.write("Please try again.")
                    # Generate a new authorization URL
                    auth_url, _ = flow.authorization_url(prompt='consent')
                    st.markdown(f"[Click here to authenticate with Google]({auth_url})")
                    st.stop()
            else:
                # No code parameter, start the auth flow
                auth_url, _ = flow.authorization_url(prompt='consent', access_type='offline')
                st.warning("You need to authenticate with Google to access your emails.")
                st.markdown(f"[Click here to authenticate with Google]({auth_url})")
                st.stop()
        except Exception as e:
            st.error(f"Error during authentication setup: {str(e)}")
            st.write("Please check your configuration and try again.")
            st.stop()

    try:
        # Build the Gmail service with our credentials
        service = build('gmail', 'v1', credentials=st.session_state.gmail_creds)
        return service
    except Exception as e:
        st.error(f"Error building Gmail service: {str(e)}")
        # Clear credentials to restart auth flow
        if "gmail_creds" in st.session_state:
            del st.session_state.gmail_creds
        st.stop()

def fetch_recent_emails(service, start_datetime, end_datetime, desired_count=50, max_fetch=200):
    """
    Fetch emails by scanning up to max_fetch messages and returning those whose internalDate
    falls between start_datetime and end_datetime.
    """
    matching_emails = []
    next_page_token = None
    fetched_messages = 0

    progress_bar = st.progress(0)
    status_text = st.empty()
    status_text.write("Fetching emails...")
    
    try:
        while fetched_messages < max_fetch and len(matching_emails) < desired_count:
            params = {"userId": "me", "maxResults": 50}
            if next_page_token:
                params["pageToken"] = next_page_token
            
            result = service.users().messages().list(**params).execute()
            messages = result.get("messages", [])
            fetched_messages += len(messages)
            
            if not messages:
                break
                
            for i, msg in enumerate(messages):
                progress_bar.progress(min(1.0, fetched_messages / max_fetch))
                status_text.write(f"Fetching emails... ({fetched_messages} scanned, {len(matching_emails)} matched)")
                
                try:
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
                except Exception as e:
                    st.warning(f"Error processing an email: {str(e)}")
                    continue
                    
            next_page_token = result.get("nextPageToken")
            if not next_page_token:
                break
                
        status_text.write(f"Fetched {len(matching_emails)} emails in the selected range out of {fetched_messages} messages scanned.")
        return matching_emails
        
    except Exception as e:
        status_text.error(f"Error fetching emails: {str(e)}")
        return []



# Then update the analyze_email_openai function to better handle missing API key
def analyze_email_openai(email_text: str) -> str:
    """
    Uses OpenAI's ChatCompletion to extract structured info from an email.
    """
    if not openai.api_key:
        st.error("OpenAI API key not configured. Please add it to your secrets.toml file.")
        return "API key missing - unable to analyze email content."
    
    prompt = (
        f"Extract all relevant details for a service request from the email below:\n\n"
        f"{email_text}\n\n"
        "Return the extracted information in a clear, structured format."
    )
    
    max_retries = 3
    wait_time = 2  # seconds
    
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
        except Exception as e:
            if attempt < max_retries - 1:
                st.warning(f"OpenAI API error: {str(e)}. Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
                wait_time *= 2  # Exponential backoff
            else:
                st.error(f"Failed to analyze email after {max_retries} attempts: {str(e)}")
                return f"Error analyzing email: {str(e)}"

def compute_priority(email, extracted_info):
    # Calculate urgency score by counting urgency-related keywords in the extracted info.
    urgency_keywords = ['urgent', 'asap', 'immediately', 'critical', 'emergency', 'deadline', 'rush']
    urgency_score = sum(1 for word in urgency_keywords if word in extracted_info.lower())
    
    # If the email is from PWC, give it a high bonus to ensure it is prioritized.
    from_field = email["from"].lower()
    bonus = 0
    
    # Check for high-priority senders
    if "pwc" in from_field:
        bonus = 100
    elif any(domain in from_field for domain in ["@acme.com", "@important-client.com"]):
        bonus = 50
        
    # Check for urgency in subject
    subject_urgency = sum(1 for word in urgency_keywords if word in email["subject"].lower())
    
    total_score = bonus + (urgency_score * 2) + (subject_urgency * 3)
    return total_score


def main():
    st.set_page_config(
        page_title="Email Summary Dashboard",
        page_icon="📧",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.title("📧 Overnight Email Summary Dashboard")
    
    # Add a sidebar with app info
    with st.sidebar:
        st.header("About")
        st.write(
            "This app analyzes your recent emails and prioritizes them based on content. "
            "It uses OpenAI to extract key information and ranks emails by urgency."
        )
        st.write("---")
        st.write("Made with ❤️ by PrezLab")
        
        # Add logout option
        if "gmail_creds" in st.session_state:
            if st.button("Logout from Google"):
                del st.session_state.gmail_creds
                st.success("Logged out successfully!")
                st.rerun()
    
    # Default date/time: yesterday at 5:00 PM to today at 9:00 AM
    today = datetime.today().date()
    yesterday = today - timedelta(days=1)
    default_start_time = dt_time(17, 0)  # 5:00 PM
    default_end_time = dt_time(9, 0)     # 9:00 AM

    # Date range selector
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Start")
        start_date = st.date_input("Date", value=yesterday, key="start_date")
        start_time = st.time_input("Time", value=default_start_time, key="start_time")
    with col2:
        st.subheader("End")
        end_date = st.date_input("Date", value=today, key="end_date")
        end_time = st.time_input("Time", value=default_end_time, key="end_time")

    # Combine date and time into datetime objects (local time)
    start_datetime = datetime.combine(start_date, start_time)
    end_datetime = datetime.combine(end_date, end_time)

    st.write("Analyzing emails from:", start_datetime, "to", end_datetime)

    # Options for analysis
    max_emails = st.slider("Maximum emails to analyze", min_value=5, max_value=100, value=20)
    
    if st.button("Run Analysis", type="primary"):
        try:
            # Get Gmail service with proper authentication
            with st.spinner("Authenticating with Google..."):
                service = get_gmail_service()
            
            # Fetch emails for the specified period
            emails = fetch_recent_emails(
                service, 
                start_datetime, 
                end_datetime, 
                desired_count=max_emails, 
                max_fetch=max_emails * 5
            )

            if not emails:
                st.info("No emails found in the selected date/time range.")
                return

            # Process the emails with OpenAI
            processed_emails = []
            progress_bar = st.progress(0)
            analysis_status = st.empty()
            
            for i, email in enumerate(emails):
                progress = (i + 1) / len(emails)
                progress_bar.progress(progress)
                analysis_status.write(f"Processing email {i+1} of {len(emails)}: {email['subject']}")
                
                analysis = analyze_email_openai(email["snippet"])
                priority = compute_priority(email, analysis)
                
                email.update({
                    "analysis": analysis,
                    "priority": priority,
                })
                processed_emails.append(email)
            
            # Sort emails by priority
            sorted_emails = sorted(processed_emails, key=lambda x: x["priority"], reverse=True)
            
            # Clear progress indicators
            progress_bar.empty()
            analysis_status.empty()
            
            # Display results
            st.success(f"Analysis complete! Analyzed {len(emails)} emails.")
            
            # Display prioritized emails
            st.subheader("📋 Prioritized Email Summaries")
            
            # Create tabs for different views
            tab1, tab2 = st.tabs(["Card View", "Table View"])
            
            with tab1:
                # Card view
                for i, email in enumerate(sorted_emails):
                    with st.container():
                        col1, col2 = st.columns([1, 4])
                        with col1:
                            st.metric("Priority", email["priority"])
                        with col2:
                            st.subheader(email["subject"])
                            st.caption(f"From: {email['from']} | {email['received_at'].strftime('%Y-%m-%d %H:%M')}")
                        st.markdown(f"**Summary:** {email['analysis']}")
                        st.markdown("---")
            
            with tab2:
                # Table view
                table_data = []
                for email in sorted_emails:
                    table_data.append({
                        "Priority": email["priority"],
                        "Subject": email["subject"],
                        "From": email["from"],
                        "Date": email["received_at"].strftime("%Y-%m-%d %H:%M"),
                        "Summary": email["analysis"]
                    })
                st.dataframe(table_data, use_container_width=True)
                
        except Exception as e:
            st.error(f"An error occurred during analysis: {str(e)}")


if __name__ == '__main__':
    main()