from flask import Flask, request, jsonify, session, redirect
from flask_cors import CORS
import requests
import secrets
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient, BlobClient, ContainerClient
import os
import json
from datetime import datetime
import base64
import uuid
from dotenv import load_dotenv
load_dotenv()

FRONTEND = os.getenv('FRONTEND')

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)
cors = CORS(app, resources={r"/*": {"origins": FRONTEND}}, supports_credentials=True)

# Microsoft App Details
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
REDIRECT_URI = os.getenv('REDIRECT_URI')
AUTH_URL = os.getenv('AUTH_URL')
TOKEN_URL = os.getenv('TOKEN_URL')



# Blob Storage Setup

CONNECTION_STRING = os.getenv('CONNECTION_STRING')
blob_service_client = BlobServiceClient.from_connection_string(CONNECTION_STRING)

# First container -- Storing sent mails data
container_name1 = "sendmaildata"
container_client1 = blob_service_client.get_container_client(container_name1)

# Second container -- Storing Recieved mails data
container_name2 = "receivedmaildata"
container_client2 = blob_service_client.get_container_client(container_name2)


@app.after_request
def after_request(response):
    response.headers.add('Access-Control-Allow-Origin', FRONTEND)
    response.headers.add('Access-Control-Allow-Credentials', 'true')
    response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,PATCH,OPTIONS')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type, Authorization')
    return response



@app.route('/login')
def login():
    
    auth_url = f"{AUTH_URL}?client_id={CLIENT_ID}&response_type=code&redirect_uri={REDIRECT_URI}&scope=Mail.Read Mail.Send User.Read"
    return redirect(auth_url)

@app.route('/get_access_token', methods=['POST'])
def get_access_token():
    try:
        data = request.json
        grant_type = data.get("grant_type")
        code_or_token = data.get("code") or data.get("refresh_token")

        if not grant_type or not code_or_token:
            return jsonify({"error": "Missing grant_type or code/refresh_token"}), 400

        token_data = {
            'client_id': CLIENT_ID,
            'client_secret': CLIENT_SECRET,
            'redirect_uri': REDIRECT_URI,
            'grant_type': grant_type,
        }

        if grant_type == "authorization_code":
            token_data["code"] = code_or_token
        elif grant_type == "refresh_token":
            token_data["refresh_token"] = code_or_token
        else:
            return jsonify({"error": "Invalid grant_type provided"}), 400

        response = requests.post(TOKEN_URL, data=token_data)

        if response.status_code == 200:
            token_response = response.json()
            return jsonify({
                "access_token": token_response.get("access_token"),
                "refresh_token": token_response.get("refresh_token"),
                "expires_in": token_response.get("expires_in")
            }), 200
        else:
            return jsonify({"error": "Failed to fetch token", "details": response.json()}), response.status_code
    except Exception as e:
        return jsonify({"error": str(e)}), 500

 
def refresh_access_token():
    refresh_token = session.get('refresh_token')
    if not refresh_token:
        return None
    token_data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'redirect_uri': REDIRECT_URI
    }
    response = requests.post(TOKEN_URL, data=token_data)
    if response.status_code == 200:
        token_response = response.json()
        session['access_token'] = token_response.get("access_token")
        session['refresh_token'] = token_response.get("refresh_token")
        session['expires_in'] = token_response.get("expires_in")
        return token_response.get("access_token")
    return None



@app.route('/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({"message": "Logged out successfully"})



@app.route('/home', methods=['GET'])
def get_mails():
    token = request.args.get('token')
    print(token)
    if not token:
        return jsonify({"error": "Unauthorized"}), 401
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(
        'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages',
        headers=headers
    ) 
    if response.status_code != 200:
        return jsonify({"error": response.json()}), response.status_code

    
    emails_data = response.json()
    print(emails_data)

    
    timestamp = datetime.now()
    filename = f"emails_{timestamp}.json"

    
    blob_client = container_client2.get_blob_client(filename)
    blob_client.upload_blob(json.dumps(emails_data), overwrite=True)

    
    return jsonify(emails_data)
   

@app.route('/send-email', methods=['POST'])
def send_email():
    try:
        # Extract email details from the request
        data = request.json
        access_token = request.headers.get('Authorization').replace("Bearer ", "")
        to_email = data.get("to_email")
        subject = data.get("subject")
        body = data.get("body")
        attachments = data.get("attachments", [])  # Attachments is an empty list by default

        if not access_token:
            return jsonify({"error": "Access token is required"}), 401

        if not to_email or not subject or not body:
            return jsonify({"error": "Recipient, subject, and body are required"}), 400

        # Split the 'to_email' by commas and remove extra spaces
        to_email_list = [email.strip() for email in to_email.split(',')]

        # Prepare the email payload
        email_payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": body
                },
                "toRecipients": [
                    {"emailAddress": {"address": email}} for email in to_email_list
                ],
                "attachments": []
            }
        }

        # Add attachments if available
        if attachments:
            for attachment in attachments:
                # Ensure file_data is a valid base64 string and decode it
                if 'file_data' in attachment:
                    file_data = attachment['file_data']
                    
                    # If the file_data is base64 encoded with a prefix (data URL), strip it
                    if file_data.startswith("data:"):
                        file_data = file_data.split(',')[1]
                    
                    # Decode the base64 string into bytes
                    file_bytes = base64.b64decode(file_data)
                    
                    email_payload['message']['attachments'].append({
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": attachment['filename'],
                        "contentBytes": base64.b64encode(file_bytes).decode("utf-8")
                    })

        # Set the headers for authorization
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }

        # Send the email via Microsoft Graph API
        send_email_url = "https://graph.microsoft.com/v1.0/me/sendMail"
        response = requests.post(send_email_url, headers=headers, json=email_payload)

        if response.status_code == 202:
            # Prepare the email data to be saved in Blob Storage
            email_data = {
                "to_email": to_email_list,
                "subject": subject,
                "body": body,
                "attachments": attachments,
                "status": "Sent",
                "timestamp": datetime.utcnow().isoformat()
            }

            # Store the email data in Blob Storage as a JSON file
            blob_name = f"sent_email_{datetime.now()}.json"
            blob_client = container_client1.get_blob_client(blob_name)
            blob_client.upload_blob(json.dumps(email_data), overwrite=True)

            return jsonify({"message": "Email sent successfully and stored in Blob!"}), 202
        else:
            return jsonify({"error": response.json()}), response.status_code

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    
# ---------------- Teams ----------------->
BLOB_CONTAINER_NAME = 'teamsdata'
BLOB_NAME = 'teams.json'

# Initialize Azure Blob Service Client
blob_service_client = BlobServiceClient.from_connection_string(CONNECTION_STRING)
container_client = blob_service_client.get_container_client(BLOB_CONTAINER_NAME)

# Ensure the container exists
try:
    container_client.get_container_properties()
except Exception:
    container_client = blob_service_client.create_container(BLOB_CONTAINER_NAME)

# Helper function to get teams data from Azure Blob Storage
def get_teams_data():
    try:
        blob_client = container_client.get_blob_client(BLOB_NAME)
        stream = blob_client.download_blob()
        data = stream.readall()
        teams = json.loads(data)
    except Exception:
        teams = {}
    return teams

# Helper function to save teams data to Azure Blob Storage
def save_teams_data(teams):
    blob_client = container_client.get_blob_client(BLOB_NAME)
    data = json.dumps(teams)
    blob_client.upload_blob(data, overwrite=True)

# Route to get all teams
@app.route('/teams', methods=['GET'])
def get_teams():
    teams = get_teams_data()
    return jsonify(list(teams.values())), 200

# Route to get a specific team by ID
@app.route('/teams/<team_id>', methods=['GET'])
def get_team(team_id):
    teams = get_teams_data()
    team = teams.get(team_id)
    if team:
        return jsonify(team), 200
    else:
        return jsonify({'error': 'Team not found'}), 404

# Route to create a new team (now accepts a description)
@app.route('/teams', methods=['POST'])
def create_team():
    data = request.get_json()
    if not data or 'name' not in data:
        return jsonify({'error': 'Team name is required'}), 400

    teams = get_teams_data()
    team_id = str(uuid.uuid4())
    team = {
        'id': team_id,
        'name': data['name'],
        'description': data.get('description', ''),
        'members': []
    }
    teams[team_id] = team
    save_teams_data(teams)
    return jsonify(team), 201

# Route to update a team by ID (optionally update name, description, or members)
@app.route('/teams/<team_id>', methods=['PUT'])
def update_team(team_id):
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No data provided'}), 400

    teams = get_teams_data()
    team = teams.get(team_id)
    if not team:
        return jsonify({'error': 'Team not found'}), 404

    if 'name' in data:
        team['name'] = data['name']
    if 'description' in data:
        team['description'] = data['description']
    if 'members' in data:
        team['members'] = data['members']

    teams[team_id] = team
    save_teams_data(teams)
    return jsonify(team), 200

# Route to delete a team by ID
@app.route('/teams/<team_id>', methods=['DELETE'])
def delete_team(team_id):
    teams = get_teams_data()
    if team_id in teams:
        del teams[team_id]
        save_teams_data(teams)
        return jsonify({'message': 'Team deleted'}), 200
    else:
        return jsonify({'error': 'Team not found'}), 404

# Route to add a member to a team (now requires both name and email)
@app.route('/teams/<team_id>/members', methods=['POST'])
def add_team_member(team_id):
    data = request.get_json()
    if not data or 'name' not in data or 'email' not in data:
        return jsonify({'error': 'Member name and email are required'}), 400

    teams = get_teams_data()
    team = teams.get(team_id)
    if not team:
        return jsonify({'error': 'Team not found'}), 404

    member_id = str(uuid.uuid4())
    member = {
        'id': member_id,
        'name': data['name'],
        'email': data['email']
    }
    team['members'].append(member)
    teams[team_id] = team
    save_teams_data(teams)
    return jsonify(member), 201

# Route to delete a member from a team
@app.route('/teams/<team_id>/members/<member_id>', methods=['DELETE'])
def delete_team_member(team_id, member_id):
    teams = get_teams_data()
    team = teams.get(team_id)
    if not team:
        return jsonify({'error': 'Team not found'}), 404

    members = team.get('members', [])
    updated_members = [m for m in members if m['id'] != member_id]
    if len(updated_members) == len(members):
        return jsonify({'error': 'Member not found'}), 404

    team['members'] = updated_members
    teams[team_id] = team
    save_teams_data(teams)
    return jsonify({'message': 'Member removed'}), 200
    
#-------------- sent mail -------------->
@app.route('/sent-mails')
def Sent_mails():
    token = request.args.get('token')
    if not token:
        return jsonify({"error": "Unauthorized"}), 401

    headers = {"Authorization": f"Bearer {token}"}
    
    url = 'https://graph.microsoft.com/v1.0/me/mailFolders/SentItems/messages/delta'
    
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return jsonify(response.json())
    else:
        return jsonify({"error": "Failed to fetch emails", "details": response.json()}), response.status_code

@app.route("/oAuth_redirect")
def redirec():
    code = request.args.get('code')
    if not code:
        return jsonify({"error": "Unauthorized"}), 401
        
    return redirect(f"{FRONTEND}/login?code={code}")



if __name__ == '__main__':
    app.run(host="0.0.0.0",port=8000)