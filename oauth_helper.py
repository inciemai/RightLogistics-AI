#!/usr/bin/env python3
"""
OAuth Helper Script for Microsoft Graph API Authentication

This script helps process OAuth redirect URLs and extract authorization codes
for Microsoft Graph API authentication.
"""

import sys
from urllib.parse import urlparse, parse_qs
from loguru import logger


def extract_auth_code_from_url(redirect_url: str) -> tuple:
    """
    Extract authorization code and other parameters from OAuth redirect URL.
    
    Args:
        redirect_url (str): The full redirect URL from OAuth response
        
    Returns:
        tuple: (auth_code, state, session_state, client_info) or (None, None, None, None) if failed
    """
    try:
        if not redirect_url:
            return None, None, None, None
            
        if 'code=' not in redirect_url:
            logger.error("No authorization code found in URL")
            return None, None, None, None
            
        parsed_url = urlparse(redirect_url)
        query_params = parse_qs(parsed_url.query)
        
        # Extract all relevant parameters
        auth_code = query_params.get('code', [None])[0]
        state = query_params.get('state', [None])[0]
        session_state = query_params.get('session_state', [None])[0]
        client_info = query_params.get('client_info', [None])[0]
        
        return auth_code, state, session_state, client_info
        
    except Exception as e:
        logger.error(f"Error parsing OAuth redirect URL: {str(e)}")
        return None, None, None, None


def validate_oauth_url(redirect_url: str) -> bool:
    """
    Validate if the URL is a proper Microsoft OAuth redirect URL.
    
    Args:
        redirect_url (str): The redirect URL to validate
        
    Returns:
        bool: True if valid, False otherwise
    """
    try:
        parsed_url = urlparse(redirect_url)
        
        # Check if it's a Microsoft OAuth redirect
        expected_domains = [
            'login.microsoftonline.com',
            'login.live.com',
            'account.live.com'
        ]
        
        if parsed_url.hostname not in expected_domains:
            logger.warning(f"URL domain '{parsed_url.hostname}' is not a recognized Microsoft OAuth domain")
            return False
            
        # Check if it contains required parameters
        query_params = parse_qs(parsed_url.query)
        if 'code' not in query_params:
            logger.error("URL does not contain authorization code")
            return False
            
        return True
        
    except Exception as e:
        logger.error(f"Error validating OAuth URL: {str(e)}")
        return False


def main():
    """Main function for interactive OAuth URL processing."""
    print("=" * 80)
    print("Microsoft Graph OAuth URL Processor")
    print("=" * 80)
    print()
    
    if len(sys.argv) > 1:
        # URL provided as command line argument
        redirect_url = sys.argv[1]
        print(f"Processing URL from command line argument...")
    else:
        # Interactive mode
        print("Paste your OAuth redirect URL below:")
        redirect_url = input("URL: ").strip()
    
    if not redirect_url:
        print("❌ No URL provided")
        sys.exit(1)
    
    print(f"\n📋 Processing URL: {redirect_url[:100]}{'...' if len(redirect_url) > 100 else ''}")
    
    # Validate URL
    if not validate_oauth_url(redirect_url):
        print("❌ Invalid OAuth redirect URL")
        sys.exit(1)
    
    # Extract parameters
    auth_code, state, session_state, client_info = extract_auth_code_from_url(redirect_url)
    
    if auth_code:
        print("\n✅ Successfully extracted OAuth parameters:")
        print(f"  📝 Authorization Code: {auth_code[:20]}...{auth_code[-10:] if len(auth_code) > 30 else auth_code}")
        if state:
            print(f"  🔑 State: {state}")
        if session_state:
            print(f"  🔐 Session State: {session_state}")
        if client_info:
            print(f"  ℹ️  Client Info: {client_info[:50]}{'...' if len(client_info) > 50 else ''}")
        
        print(f"\n🎯 Full Authorization Code:")
        print(f"  {auth_code}")
        
        print(f"\n💡 You can now use this authorization code in your application.")
        print(f"   Remember that authorization codes typically expire within 10 minutes.")
        
    else:
        print("❌ Failed to extract authorization code from URL")
        sys.exit(1)


if __name__ == "__main__":
    main() 