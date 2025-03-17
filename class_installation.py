import os
import os.path
import json
import requests
import logging

# Configure logging (outside the class)
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Class definition starts here
class Installation:

    def __init__(self):
        self.cwd = os.getcwd()
        self.client = None
        self.refresh_token = None

    def get_client(self, v=False):
        if self.client is None:
            self.read_client(v)
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

    def get_access_token(self, v=False):
        try:
            client = self.get_client()
            refresh_token = self.get_refresh_token()
            logger.debug("Requesting access token from %s", client['token_endpoint'])
            r = requests.post(client['token_endpoint'], data={
                'client_id': client['client_id'],
                'client_secret': client['client_secret'],
                'refresh_token': refresh_token,
                'grant_type': 'refresh_token',
                'scope': 'https://graph.microsoft.com/.default'
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