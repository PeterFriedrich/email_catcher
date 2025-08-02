"""
Email Catcher - Outlook Client Module

Handles Graph API authentication and operations
"""

import os
import requests
import time
from typing import Dict
from dotenv import load_dotenv

import traceback

load_dotenv()


class GraphAuthenticator:
    """
    Handles authentication, using device code flow.

    Manages obtaining, storing, and refreshing tokens
    needed for Microsoft Graph API.
    """
    def __init__(self):
        self.client_id = os.getenv('CLIENT_ID')
        self.tenant_id = os.getenv('TENANT_ID')
        self.access_token = None

    def authenticate(self):
        """
        Authenticate using device code flow
        makes a request to microsoft.
        Return bool, true for success, else false.
        """

        # make request urls
        base_url = 'https://login.microsoftonline.com'

        device_code_url = f"{base_url}/{self.tenant_id}/oauth2/v2.0/devicecode"
        token_url = f"{base_url}/{self.tenant_id}/oauth2/v2.0/token"

        # request device code, verification info
        read_val = 'https://graph.microsoft.com/Mail.Read'
        write_val = 'https://graph.microsoft.com/Mail.ReadWrite'

        device_data = {
            'client_id': self.client_id,
            'scope': read_val + " " + write_val
            }

        try:
            device_response = requests.post(device_code_url, data=device_data)
            device_response.raise_for_status()
            device_info = device_response.json()

            # prompt user to complete browser authentication
            print("\n Authentication required:")
            print(f"Go to: {device_info['verification_uri']}")
            print(f"Enter code: {device_info['user_code']}")
            print(f"Code expires in {device_info['expires_in']} seconds")

            # poll for completion
            grant_type_val = 'urn:ietf:params:oauth:grant-type:device_code'

            poll_data = {
                    'grant_type': grant_type_val,
                    'client_id': self.client_id,
                    'device_code': device_info['device_code']
                    }

            interval = device_info.get('interval', 5)
            expires_in = device_info['expires_in']

            # request and handle
            print("------------------------")
            print(f"interval:{interval}, expires_in:{expires_in}")
            for _ in range(0, expires_in, interval):

                time.sleep(interval)
                token_response = requests.post(token_url, data=poll_data)

                if token_response.status_code == 200:
                    token_data = token_response.json()
                    self.access_token = token_data['access_token']
                    print("Authentication successful!")
                    return True

                elif token_response.status_code == 400:
                    error = token_response.json().get('error')
                    if error == "authorization_pending":
                        print("still waiting...")
                        continue  # try again
                    elif error == "authorization_declined":
                        print("Auth declined by user")
                        return False
                    elif error == "expired_token":
                        print("Code expired")
                        return False
                elif token_response.status_code == 401:
                    print("401 Unauth")
                    print(f"Response: {token_response.text}")
                    return False

                else:
                    print("Error: printing code")
                    print(token_response.status_code)
                    print(f"Response: {token_response.text}")
                    return False

                print("Authentication timed out")
                return False

        except Exception as e:
            print(f"Authentication error: {e}")
            return False

    def get_headers(self):
        return {
                'Authorization': f'Bearer {self.access_token}',
                'Content-type': 'applications/json'
                }


class GraphAPIClient:
    """
    Manages requests to the Microsoft Graph API
    """
    def __init__(
        self,
        authenticator: GraphAuthenticator,
        base_url: str = "https://graph.microsoft.com/v1.0"
    ):
        self.base_url = base_url
        self.authenticator = auth

    def get(self, endpoint: str) -> Dict:
        """Make a GET request to Graph API."""
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        headers = self.authenticator.get_headers()

        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()


class OutlookMailClient:
    """
    Performs email specific operations via Graph API
    """

    def __init__(self, api_client: GraphAPIClient):
        self.api_client = api_client
        self._junk_folder_id = None

    def get_mail_folders(self):
        """
        Get all mail folders from the user.

        Returns: List of folder objects
        """
        mail_folder_url = "/me/mailFolders"
        response = self.api_client.get(mail_folder_url)
        return response.get('value', [])


# basic test
if __name__ == "__main__":
    print("=== Email Catcher Test ===")

    # auth test
    auth = GraphAuthenticator()
    if auth.authenticate():
        print("Authentication successful!")

        # test api client
        api_client = GraphAPIClient(auth)
        try:
            profile = api_client.get("/me")
            print(f"Api working, hello {profile.get('displayName')}")
        except Exception as e:
            print(f"API test failed: {e}")
    else:
        print("Authentication failed")

    # outlook client test
    email_address = os.getenv("EMAIL_ADDRESS")
    outlook = OutlookMailClient(api_client)

    try:
        folders = outlook.get_mail_folders()
        print(f"Found {len(folders)} mail folders:")
        for folder in folders:
            print(f" - {folder['displayName']} (ID: {folder['id']})")
    except Exception as e:
        print(f"Failed to get folders: {e}")
        traceback.print_exc()
