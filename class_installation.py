import os
import os.path
import json
import requests
import logging
import webbrowser

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class Installation:
    def __init__(self):
        self.cwd = os.getcwd()
        self.client = None
        self.refresh_token = None
        # Set redirect URI here; update this if different in Azure AD
        self.redirect_uri = 'http://localhost'

    def get_client(self, v=False):
        if self.client is None:
            self.read_client(v)
        if self.client is None:
            logger.error("Client is still None after read_client")
            raise ValueError("Failed to initialize client")
        return self.client

    def read_client(self, v=False):
        client_name = 'client_secret.json'
        client_path = os.path.join(self.cwd, client_name)
        try:
            with open(client_path, 'r') as f:
                client = json.load(f)
            if not client:
                raise ValueError("Client data is empty")
            required_keys = ['authorization_endpoint', 'token_endpoint', 'client_id', 'client_secret']
            missing_keys = [key for key in required_keys if key not in client]
            if missing_keys:
                raise KeyError(f"Missing required keys in client_secret.json: {missing_keys}")
            if v:
                logger.info('=== authorization endpoint ===\n%s', client['authorization_endpoint'])
                logger.info('=== token endpoint ===\n%s', client['token_endpoint'])
                logger.info('=== client id ===\n%s', client['client_id'])
                logger.info('=== client secret ===\n%s', client['client_secret'])
            self.client = client
        except FileNotFoundError:
            logger.error("Client secret file not found: %s", client_path)
            raise
        except json.JSONDecodeError as e:
            logger.error("Failed to parse JSON in %s: %s", client_path, str(e))
            raise
        except (ValueError, KeyError) as e:
            logger.error("Validation error in client data: %s", str(e))
            raise
        except Exception as e:
            logger.error("Unexpected error reading client secret: %s", str(e))
            raise

    def get_refresh_token(self, v=False):
        if self.refresh_token is None:
            self.read_refresh_token(v)
        if self.refresh_token is None:
            logger.error("Refresh token is still None after read_refresh_token")
            raise ValueError("Failed to initialize refresh token")
        return self.refresh_token

    def read_refresh_token(self, v=False):
        refresh_name = 'refresh_token.txt'
        refresh_path = os.path.join(self.cwd, refresh_name)
        try:
            with open(refresh_path, mode='r') as f:
                refresh_token = f.read().rstrip()
            if not refresh_token:
                raise ValueError("Refresh token is empty")
            if v:
                logger.info('=== refresh token ===\n%s', refresh_token)
            self.refresh_token = refresh_token
        except FileNotFoundError:
            logger.error("Refresh token file not found: %s", refresh_path)
            raise
        except Exception as e:
            logger.error("Unexpected error reading refresh token: %s", str(e))
            raise

    def set_refresh_token(self, refresh_token):
        refresh_name = 'refresh_token.txt'
        refresh_path = os.path.join(self.cwd, refresh_name)
        try:
            with open(refresh_path, mode='w') as f:
                f.write(refresh_token + '\n')
            self.refresh_token = refresh_token
        except Exception as e:
            logger.error("Failed to write refresh token to %s: %s", refresh_path, str(e))
            raise

    def get_authorization_code(self, v=False):
        try:
            client = self.get_client(v)
            scopes = [
                "https://graph.microsoft.com/Files.ReadWrite",
                "https://graph.microsoft.com/Sites.ReadWrite.All"
            ]
            auth_url = (
                f"{client['authorization_endpoint']}?"
                f"client_id={client['client_id']}&"
                f"response_type=code&"
                f"redirect_uri={self.redirect_uri}&"
                f"scope={' '.join(scopes)}"
            )
            logger.info("Opening browser for authorization: %s", auth_url)
            logger.info("Please sign in and paste the code from the URL (after 'code=').")
            webbrowser.open(auth_url)
            auth_code = input("Enter the authorization code from the browser URL: ")
            if not auth_code:
                raise ValueError("No authorization code provided")
            return auth_code
        except Exception as e:
            logger.error("Failed to get authorization code: %s", str(e))
            raise

    def get_initial_token(self, v=False):
        try:
            client = self.get_client(v)
            auth_code = self.get_authorization_code(v)
            scopes = [
                "https://graph.microsoft.com/Files.ReadWrite",
                "https://graph.microsoft.com/Sites.ReadWrite.All"
            ]
            logger.debug("Requesting initial token from %s", client['token_endpoint'])
            r = requests.post(client['token_endpoint'], data={
                'client_id': client['client_id'],
                'client_secret': client['client_secret'],
                'code': auth_code,
                'grant_type': 'authorization_code',
                'redirect_uri': self.redirect_uri,
                'scope': ' '.join(scopes)
            })
            r.raise_for_status()
            response = r.json()
            if 'error' in response:
                logger.error("Authentication failed: %s", response['error_description'])
                raise AssertionError(response['error_description'])
            self.set_refresh_token(response['refresh_token'])
            if v:
                logger.info('=== initial access token ===\n%s', response['access_token'])
                logger.info('=== new refresh token ===\n%s', response['refresh_token'])
            return response['access_token']
        except requests.RequestException as e:
            logger.error("HTTP request failed: %s", str(e))
            raise
        except Exception as e:
            logger.error("Unexpected error in get_initial_token: %s", str(e))
            raise

    def get_access_token(self, v=False):
        try:
            client = self.get_client(v)
            refresh_token = self.get_refresh_token(v)
            scopes = [
                "https://graph.microsoft.com/Files.ReadWrite",
                "https://graph.microsoft.com/Sites.ReadWrite.All"
            ]
            logger.debug("Requesting access token from %s with scopes: %s", client['token_endpoint'], scopes)
            r = requests.post(client['token_endpoint'], data={
                'client_id': client['client_id'],
                'client_secret': client['client_secret'],
                'refresh_token': refresh_token,
                'grant_type': 'refresh_token',
                'scope': ' '.join(scopes)
            })
            logger.debug("Response status: %s", r.status_code)
            logger.debug("Response body: %s", r.text)
            r.raise_for_status()
            response = r.json()
            if 'error' in response:
                logger.error("Authentication failed: %s", response['error_description'])
                raise AssertionError(response['error_description'])
            self.set_refresh_token(response['refresh_token'])
            access_token = response['access_token']
            if v:
                logger.info('=== access token ===\n%s', access_token)
            logger.debug("Access token retrieved successfully")
            return access_token
        except requests.RequestException as e:
            logger.error("HTTP request failed: %s", str(e))
            raise
        except KeyError as e:
            logger.error("Missing expected key in response: %s", str(e))
            raise
        except Exception as e:
            logger.error("Unexpected error in get_access_token: %s", str(e))
            raise
