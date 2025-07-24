"""
Email Catcher - Outlook Client Module

Handles Graph API authentication and operations
"""

import os

class GraphAuthenticator:
    def __init__(self):
        self.client_id = os.getenv('CLIENT_ID')
        self.tenant_id = os.getenv('TENANT_ID')

    def authenticate(self):
        pass

# basic test
if __name__ == "__main__":
    auth = GraphAuthenticator()
    print(auth.client_id)
    if auth.authenticate():
        print("Authentication successful!")
    else:
        print("Authentication failed")


