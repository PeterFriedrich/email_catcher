import os
from dotenv import load_dotenv

load_dotenv()

from outlook_client import GraphAuthenticator

if __name__ == "__main__":
    print("Starting email-catcher...")
    auth = GraphAuthenticator()
    # for testing of load env vars
    print(f"Client ID loaded: {auth.client_id is not None}")
