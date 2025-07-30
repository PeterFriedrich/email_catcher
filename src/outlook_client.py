"""
Email Catcher - Outlook Client Module

Handles Graph API authentication and operations
"""

import os
import requests
import time

class GraphAuthenticator:
    def __init__(self):
        # just creates
        self.client_id = os.getenv('CLIENT_ID')
        self.tenant_id = os.getenv('TENANT_ID')

    def authenticate(self):
        """
        Authenticate using device code flow
        makes a request to microsoft.
        Return bool, true for success, else false.
        """

        # make request urls
        device_code_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/devicecode"
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        # request device code, verification info
        device_data = {
            'client_id': self.client_id,
            'scope':'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.ReadWrite'
        }
        
        try:
            device_response = requests.requests.post(device_code_url, data=device_data)
            device_response.raise_for_status()
            device_info = device_response.json()
            
            # prompt user to complete browser authentication
            print(f"\n Authentication required:")
            print(f"Go to: {device_info['verification_uri']}")
            print(f"Enter code: {device_info['user_code']}")
            print(f"Code expires in {device_info['expires_in']} seconds")

        except Exception as e:
            print(f"Authentication error: {e}")
            return False

        return False

# basic test
if __name__ == "__main__":
    auth = GraphAuthenticator()
    print(auth.client_id)
    if auth.authenticate():
        print("Authentication successful!")
    else:
        print("Authentication failed")
