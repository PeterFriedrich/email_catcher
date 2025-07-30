"""
Email Catcher - Outlook Client Module

Handles Graph API authentication and operations
"""

import os
import requests
import time


class GraphAuthenticator:
    """
    """
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
            device_response = requests.requests.post(device_code_url,
                                                     data=device_data)
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

            # check for response, handle
            for _ in range(0, expires_in, interval):
                time.sleep(interval)

                token_response = requests.post(token_url, data=poll_data)

                if token_response.status_code == 200:
                    token_data = token_response.json()
                    self.access_token = token_data['access_token']
                    print("Authentication successful!")
                    return True

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
