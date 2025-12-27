"""
Google Docs API Authentication Module

This module handles authentication with Google Docs API using service account credentials.
"""

from __future__ import annotations
import json
from typing import Optional, Dict, Any
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class GoogleDocsAuth:
    """
    Handles authentication and API client creation for Google Docs API.
    
    This class manages service account authentication and provides an authenticated
    Google Docs API client instance.
    """
    
    # Required scopes for Google Docs API
    SCOPES = [
        'https://www.googleapis.com/auth/documents',
        'https://www.googleapis.com/auth/drive.file'
    ]
    
    def __init__(self, service_account_file: str):
        """
        Initialize the authentication handler.
        
        Args:
            service_account_file (str): Path to the service account JSON file
            
        Raises:
            FileNotFoundError: If the service account file doesn't exist
            ValueError: If the service account file is invalid
        """
        self.service_account_file = Path(service_account_file)
        self._credentials: Optional[service_account.Credentials] = None
        self._service: Optional[Any] = None
        
        if not self.service_account_file.exists():
            raise FileNotFoundError(f"Service account file not found: {service_account_file}")
        
        self._validate_service_account_file()
    
    def _validate_service_account_file(self) -> None:
        """
        Validate that the service account file contains required fields.
        
        Raises:
            ValueError: If the service account file is missing required fields
        """
        try:
            with open(self.service_account_file, 'r') as f:
                service_account_info = json.load(f)
            
            required_fields = ['type', 'project_id', 'private_key_id', 'private_key', 'client_email']
            missing_fields = [field for field in required_fields if field not in service_account_info]
            
            if missing_fields:
                raise ValueError(f"Service account file missing required fields: {missing_fields}")
                
            if service_account_info.get('type') != 'service_account':
                raise ValueError("Service account file must be of type 'service_account'")
                
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in service account file: {e}")
    
    def authenticate(self) -> service_account.Credentials:
        """
        Authenticate using the service account credentials.
        
        Returns:
            service_account.Credentials: Authenticated credentials object
            
        Raises:
            Exception: If authentication fails
        """
        try:
            self._credentials = service_account.Credentials.from_service_account_file(
                str(self.service_account_file),
                scopes=self.SCOPES
            )
            
            # Refresh the credentials to ensure they're valid
            request = Request()
            self._credentials.refresh(request)
            
            return self._credentials
            
        except Exception as e:
            raise Exception(f"Authentication failed: {e}")
    
    def get_service(self) -> Any:
        """
        Get an authenticated Google Docs API service instance.
        
        Returns:
            googleapiclient.discovery.Resource: Authenticated Google Docs service
            
        Raises:
            Exception: If service creation fails
        """
        if not self._credentials:
            self.authenticate()
        
        try:
            self._service = build('docs', 'v1', credentials=self._credentials)
            return self._service
            
        except Exception as e:
            raise Exception(f"Failed to create Google Docs service: {e}")
    
    def get_drive_service(self) -> Any:
        """
        Get an authenticated Google Drive API service instance.
        
        This is useful for document creation and sharing operations.
        
        Returns:
            googleapiclient.discovery.Resource: Authenticated Google Drive service
            
        Raises:
            Exception: If service creation fails
        """
        if not self._credentials:
            self.authenticate()
        
        try:
            drive_service = build('drive', 'v3', credentials=self._credentials)
            return drive_service
            
        except Exception as e:
            raise Exception(f"Failed to create Google Drive service: {e}")
    
    @property
    def credentials(self) -> Optional[service_account.Credentials]:
        """Get the current credentials object."""
        return self._credentials
    
    @property
    def is_authenticated(self) -> bool:
        """Check if credentials are available and valid."""
        return (
            self._credentials is not None and 
            self._credentials.valid and 
            not self._credentials.expired
        )
    
    def refresh_credentials(self) -> None:
        """
        Refresh the authentication credentials.
        
        Raises:
            Exception: If credential refresh fails
        """
        if not self._credentials:
            raise Exception("No credentials to refresh. Call authenticate() first.")
        
        try:
            request = Request()
            self._credentials.refresh(request)
        except Exception as e:
            raise Exception(f"Failed to refresh credentials: {e}")


def create_authenticated_client(service_account_file: str) -> GoogleDocsAuth:
    """
    Convenience function to create and authenticate a Google Docs client.
    
    Args:
        service_account_file (str): Path to the service account JSON file
        
    Returns:
        GoogleDocsAuth: Authenticated client instance
        
    Raises:
        Exception: If authentication fails
    """
    auth_client = GoogleDocsAuth(service_account_file)
    auth_client.authenticate()
    return auth_client
