import os 
import logging 
from fastapi import FastAPI, BackgroundTasks 
from office365.sharepoint.client_context import ClientContext 
from office365.runtime.auth.user_credential import UserCredential 
from apscheduler.schedulers.background import BackgroundScheduler 
from datetime import datetime 
from dotenv import load_dotenv 
 
# Load environment variables (store credentials securely in a .env file) 
load_dotenv() 
 
SHAREPOINT_URL = os.getenv("SHAREPOINT_URL")  # Main SharePoint URL 
USERNAME = os.getenv("SHAREPOINT_USERNAME")  # SharePoint Username 
PASSWORD = os.getenv("SHAREPOINT_PASSWORD")  # SharePoint Password 
 
# List of 13 sites under the main SharePoint site 
SITES = [ 
    "Site1", "Site2", "Site3", "Site4", "Site5", "Site6", "Site7", "Site8",  
    "Site9", "Site10", "Site11", "Site12", "Site13" 
] 
 
# Setup logging 
LOG_FILE = "fetch_logs.log" 
logging.basicConfig( 
    filename=LOG_FILE, 
    level=logging.INFO, 
    format="%(asctime)s - %(levelname)s - %(message)s", 
) 
logger = logging.getLogger(__name__) 
 
# FastAPI application 
app = FastAPI() 
 
# Function to fetch documents from SharePoint and store them in folders 
def fetch_documents(): 
    try: 
        logger.info("Starting document fetch process.") 
        ctx = ClientContext(SHAREPOINT_URL).with_credentials(UserCredential(USERNAME, PASSWORD)) 
 
        for site in SITES: 
            site_url = f"{SHAREPOINT_URL}/sites/{site}" 
            ctx.site = ClientContext(site_url).with_credentials(UserCredential(USERNAME, PASSWORD)) 
            libraries = ctx.web.lists.get().execute_query() 
 
            for library in libraries: 
                if library.properties.get("BaseTemplate") != 101:  # 101 = Document Library 
                    continue  # Skip non-document libraries 
 
                folder_path = f"SharePoint_Documents/{site}/{library.properties['Title']}" 
                os.makedirs(folder_path, exist_ok=True)  # Create directory structure 
 
                files = library.root_folder.files.get().execute_query() 
                for file in files: 
                    file_name = file.properties["Name"] 
                    file_content = file.read().execute_query() 
 
                    file_path = os.path.join(folder_path, file_name) 
                    with open(file_path, "wb") as f: 
                        f.write(file_content) 
 
                    logger.info(f"Downloaded: {file_name} -> {file_path}") 
 
        logger.info("Document fetch completed successfully.") 
 
    except Exception as e: 
        logger.error(f"Error fetching documents: {str(e)}", exc_info=True) 
 
# Scheduler to run fetch_documents() every 3 months 
scheduler = BackgroundScheduler() 
scheduler.add_job(fetch_documents, "interval", months=3, next_run_time=datetime.now()) 
scheduler.start() 
 
@app.get("/") 
def home(): 
    return {"message": "SharePoint Document Fetch API is running"} 
 
@app.get("/fetch") 
def manual_fetch(background_tasks: BackgroundTasks): 
    background_tasks.add_task(fetch_documents) 
    return {"message": "Fetching started in the background"}