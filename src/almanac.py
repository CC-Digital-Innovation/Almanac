import base64
import os
import secrets
from pathlib import PurePath

import requests
from dotenv import load_dotenv
from fastapi import Depends, FastAPI, HTTPException, status
from fastapi.security import APIKeyHeader
from loguru import logger
from pydantic import BaseModel


# Load secrets from environemnt variables defined in deployement.
load_dotenv(PurePath(__file__).with_name('.env'))

# Assign environment variables to globals.
ALMANAC_API_KEY = os.getenv('ALMANAC_API_KEY')
AZURE_DEVOPS_PERSONAL_ACCESS_TOKEN = os.getenv('AZURE_DEVOPS_PERSONAL_ACCESS_TOKEN')
MICROSOFT_TEAMS_WEBHOOK_URL = os.getenv('MICROSOFT_TEAMS_WEBHOOK_URL')

# Initialize the FastAPI app.
app = FastAPI()

# Initialize the API key for authorization.
api_key = APIKeyHeader(name='API-Key-Name')

# Define a Pydantic model for the error message payload.
class ErrorMessage(BaseModel):
    text: str


# Authorize the request using the API key.
def authorize(key: str = Depends(api_key)):
    if not secrets.compare_digest(key, ALMANAC_API_KEY):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail='Invalid token'
        )


# Post error message to Microsoft Teams.
@app.post('/ms_teams', dependencies=[Depends(authorize)])
def send_alert_to_microsoft_teams(error_message: ErrorMessage):
    # Initialize the payload for Microsoft Teams Adaptive Card.
    microsoft_teams_adaptive_card_payload = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.2",
        "body": [
            {
                "type": "TextBlock",
                "text": "Error Alert from Almanac",
                "wrap": True
            },
            {
                "type": "TextBlock",
                "text": error_message.text,
                "wrap": True
            }
        ]
    }

    # Send the payload to Microsoft Teams using the webhook URL.
    try:
        response = requests.post(
            url=MICROSOFT_TEAMS_WEBHOOK_URL,
            json=microsoft_teams_adaptive_card_payload,
            headers={'Content-Type': 'application/json'}
        )
    except Exception as e:
        logger.error(f"Error sending alert to Microsoft Teams: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to send alert to Microsoft Teams: {e}"
        )
    # Check if the request was successful and log the result.
    else:
        if not response or not 200 <= response.status_code <= 299:
            logger.error(f"Failed to send alert to Microsoft Teams. Status code: {response.status_code}, Response: {response.text}")
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Failed to send alert to Microsoft Teams: {e}"
            )
        else:
            logger.info("Alert sent to Microsoft Teams successfully.")


# Post error message to Azure DevOps.
@app.post('/azure_devops', dependencies=[Depends(authorize)])
def send_alert_to_azure_devops(error_message: ErrorMessage):
    # Initialize the Azure DevOps API URL and headers for work item creation.
    organization = "teramachCC"
    project = "DevOps - Internal"
    work_item_type = "Issue"
    azure_devops_url = f"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/${work_item_type}?api-version=7.1"
    
    # Authorization header (Personal Access Token needs to be base64 encoded).
    encoded_pat = base64.b64encode(f":{AZURE_DEVOPS_PERSONAL_ACCESS_TOKEN}".encode()).decode()
    headers = {
        'Content-Type': 'application/json-patch+json',
        'Authorization': f'Basic {encoded_pat}'
    }
    
    # Initialize the payload for Azure DevOps work item creation.
    payload = [
        {
            "op": "add",
            "path": "/fields/System.Title",
            "from": None,
            "value": "Error message from Almanac"
        },
        {
            "op": "add",
            "path": "/fields/System.Description",
            "value": error_message.text
        }
    ]
    
    # Send the payload to Azure DevOps.
    try:
        response = requests.post(
            url=azure_devops_url,
            json=payload,
            headers=headers
        )
    except Exception as e:
        logger.error(f"Error sending alert to Azure DevOps: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error sending alert to Azure DevOps: {e}"
        )
    # Check if the request was successful and log the result.
    else:
        if not response or not 200 <= response.status_code <= 299:
            logger.error(f"Failed to send alert to Azure DevOps. Status code: {response.status_code}, Response: {response.text}")
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Failed to send alert to Azure DevOps: {e}"
            )
        else:
            logger.info("Alert sent to Azure DevOps successfully.")
