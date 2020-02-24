import json
from pathlib import Path
from typing import Optional


class User:
    def __init__(self, jsonPath: str):
        self.URL: Optional[str] = None
        self.SSO_ID: Optional[str] = None
        self.SSO_PASSWORD: Optional[str] = None

        if not Path(jsonPath).is_file():
            raise Exception(f"JSON File Not Found: {jsonPath}")

        with open(jsonPath, 'r') as f:
            json_data = json.load(f)
            try:
                self.URL = json_data["user"]["url"]
                self.SSO_ID = json_data["user"]["sso_id"]
                self.SSO_PASSWORD = json_data["user"]["sso_password"]
            except KeyError as e:
                raise Exception("JSON Key Not Found")


if __name__ == "__main__":
    user = User("user.json")
