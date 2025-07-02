from O365 import Account, FileSystemTokenBackend, MSGraphProtocol
from O365.message import Message
from azure.identity import ClientSecretCredential
import asyncio
from datetime import datetime, timedelta
from loguru import logger
from config import settings
from gemini_processor import GeminiProcessor
from typing import List, Dict, Optional
import os

class EmailProcessor:
    def __init__(self, user_id: Optional[str] = None):
        """
        Initialize EmailProcessor for a specific user or all configured users.
        
        Args:
            user_id (str, optional): Specific user ID to process. If None, will process all configured users.
        """
        self.user_id = user_id
        self.account = self._initialize_account()
        self.mailbox = None  # Initialize in fetch_unread_emails
        self.gemini_processor = GeminiProcessor()
        
    def _initialize_account(self):
        """Initialize O365 Account with OAuth credentials."""
        try:
            # Initialize account with client credentials
            credentials = (settings.OUTLOOK_CLIENT_ID, settings.OUTLOOK_CLIENT_SECRET)
            token_backend = FileSystemTokenBackend(token_path='.')
            
            # Create protocol with specific settings
            protocol = MSGraphProtocol()
            protocol.api_version = 'v1.0'
            
            # Initialize account with protocol
            account = Account(credentials,
                            auth_flow_type='credentials',
                            tenant_id=settings.OUTLOOK_TENANT_ID,
                            token_backend=token_backend,
                            protocol=protocol)
            
            # Ensure authentication
            if not account.is_authenticated:
                if not account.authenticate():
                    raise Exception("Failed to authenticate with Microsoft Graph API")
            return account
        except Exception as e:
            logger.error(f"Failed to initialize O365 account: {str(e)}")
            raise
            
    async def fetch_unread_emails(self, batch_size: int = None, user_id: str = None) -> list:
        """
        Fetch unread emails for a specific user using proper O365 methods.
        
        Args:
            batch_size (int, optional): Number of emails to fetch. Defaults to settings.BATCH_SIZE.
            user_id (str, optional): User ID to fetch emails for. Uses instance user_id if not provided.
            
        Returns:
            list: List of email messages
        """
        try:
            batch_size = batch_size or settings.BATCH_SIZE_PER_ACCOUNT
            target_user_id = user_id or self.user_id
            
            if not target_user_id:
                raise ValueError("User ID must be provided either during initialization or method call")
            
            # Initialize mailbox if not already done or user changed
            if not self.mailbox or (hasattr(self, '_current_user') and self._current_user != target_user_id):
                # Use the specific user's mailbox
                self.mailbox = self.account.mailbox(resource=target_user_id)
                self._current_user = target_user_id
            
            # Simplified filter - only check for unread emails
            # We'll filter out self-sent emails in the application code
            filter_query = "isRead eq false"
            query = self.mailbox.inbox_folder().get_messages(
                limit=batch_size * 2,  # Fetch more to account for filtering
                query=filter_query,
                order_by="receivedDateTime desc"
            )
            
            messages = []
            user_email = target_user_id.lower()
            
            for message in query:
                # Filter out emails sent by the user themselves
                try:
                    sender = getattr(message, 'sender', None)
                    if sender and hasattr(sender, 'address'):
                        sender_email = getattr(sender, 'address', '').lower()
                        if sender_email == user_email:
                            continue  # Skip emails sent by user themselves
                except Exception as e:
                    logger.warning(f"Could not verify sender for message: {str(e)}")
                    # Include the message if we can't verify sender
                
                messages.append(message)
                
                # Stop when we have enough messages
                if len(messages) >= batch_size:
                    break
            
            logger.info(f"Fetched {len(messages)} unread emails for user {target_user_id}")
            return messages
            
        except Exception as e:
            logger.error(f"Error fetching emails for user {target_user_id}: {str(e)}")
            # Fallback to direct API call if O365 methods fail
            return await self._fetch_emails_direct_api(batch_size, target_user_id)
            
    async def _fetch_emails_direct_api(self, batch_size: int, user_id: str) -> list:
        """Fallback method using direct API calls."""
        try:
            # Use the specific user's endpoint
            url = self.account.protocol.service_url
            messages_endpoint = f'{url}/users/{user_id}/messages'
            # Simplified filter - only check for unread emails
            filter_clause = "isRead eq false"
            params = {
                '$filter': filter_clause,
                '$orderby': 'receivedDateTime desc',
                '$top': str(batch_size * 2),  # Fetch more to account for filtering
                '$select': 'id,subject,body,from,toRecipients,receivedDateTime,isRead'
            }
            
            # Make the request using the account's connection
            response = self.account.con.get(messages_endpoint, params=params)
            response.raise_for_status()
            
            # Convert response to Message objects with proper parent context
            messages = []
            user_email = user_id.lower()
            
            for msg_data in response.json().get('value', []):
                # Filter out emails sent by the user themselves
                try:
                    from_field = msg_data.get('from', {})
                    sender_email = from_field.get('emailAddress', {}).get('address', '').lower()
                    if sender_email == user_email:
                        continue  # Skip emails sent by user themselves
                except Exception as e:
                    logger.warning(f"Could not verify sender for message: {str(e)}")
                    # Include the message if we can't verify sender
                
                # Ensure mailbox is initialized for the correct user
                if not self.mailbox or (hasattr(self, '_current_user') and self._current_user != user_id):
                    self.mailbox = self.account.mailbox(resource=user_id)
                    self._current_user = user_id
                    
                message = Message(parent=self.mailbox.inbox_folder(), **msg_data)
                messages.append(message)
                
                # Stop when we have enough messages
                if len(messages) >= batch_size:
                    break
            
            logger.info(f"Fetched {len(messages)} unread emails via direct API for user {user_id}")
            return messages
            
        except Exception as e:
            logger.error(f"Error in direct API fetch for user {user_id}: {str(e)}")
            raise

    async def process_email(self, email, user_id: str = None) -> dict:
        """Process a single email through the extraction pipeline."""
        try:
            target_user_id = user_id or self.user_id
            
            # Validate email object
            if not email:
                logger.error("Invalid email object received")
                return None
                
            # Get email ID for logging
            email_id = getattr(email, 'object_id', getattr(email, 'id', 'unknown'))

            # Additional safety check: Skip emails sent by the user themselves
            try:
                sender = getattr(email, 'sender', None)
                if sender and hasattr(sender, 'address'):
                    sender_email = getattr(sender, 'address', '').lower()
                    user_email = target_user_id.lower()
                    if sender_email == user_email:
                        logger.info(f"Skipping email {email_id} - sent by user themselves ({sender_email})")
                        return None
            except Exception as e:
                logger.warning(f"Could not verify sender for email {email_id}: {str(e)}")
                # Continue processing if we can't verify sender

            # Prepare email content
            try:
                email_content = self._prepare_email_content(email)
                if not email_content:
                    logger.error(f"Could not prepare content for email {email_id}")
                    return None
            except Exception as e:
                logger.error(f"Error preparing email content for {email_id}: {str(e)}")
                return None

            # Process with Gemini
            extracted_data = await self.gemini_processor.process_email(email_content)
            
            # Check if we got meaningful data (not just default structure)
            has_meaningful_data = (extracted_data and 
                                 any(value is not None and value != [] and value != "" 
                                     for key, value in extracted_data.items() 
                                     if key not in ['confidence_score', 'processing_timestamp', 'email_id', 'user_id']))
            
            if has_meaningful_data:
                # Add metadata
                extracted_data.update({
                    "processing_timestamp": datetime.utcnow().isoformat(),
                    "email_id": email_id,
                    "user_id": target_user_id,
                    "sender": self._get_sender_info(email),
                    "subject": getattr(email, 'subject', 'No subject'),
                    "received_datetime": getattr(email, 'received', datetime.utcnow()).isoformat() if hasattr(email, 'received') else datetime.utcnow().isoformat()
                })
                
                logger.info(f"Successfully processed email {email_id} for user {target_user_id} with confidence {extracted_data.get('confidence_score', 0)}")
                
                # Mark email as read after successful processing
                await self._mark_email_as_read(email)
                
                return extracted_data
            else:
                logger.info(f"No meaningful data extracted from email {email_id} for user {target_user_id}, skipping")
                return None
                
        except Exception as e:
            logger.error(f"Error processing email: {str(e)}")
            return None

    def _prepare_email_content(self, email) -> str:
        """Prepare email content for processing."""
        try:
            # Get basic email information
            subject = getattr(email, 'subject', 'No subject')
            body = ""
            
            # Try to get email body
            try:
                if hasattr(email, 'body'):
                    body = str(email.body)
                elif hasattr(email, 'get_body'):
                    body = email.get_body()
                else:
                    # Try different body attributes
                    for attr in ['text_body', 'html_body', 'content']:
                        if hasattr(email, attr):
                            body = str(getattr(email, attr))
                            break
            except Exception as e:
                logger.warning(f"Could not extract body: {str(e)}")
                body = "Could not extract email body"

            # Get sender information
            sender_info = self._get_sender_info(email)
            
            # Get recipient information
            recipient_info = self._get_recipient_info(email)
            
            # Get date/time information
            received_time = ""
            try:
                if hasattr(email, 'received'):
                    received_time = str(email.received)
                elif hasattr(email, 'receivedDateTime'):
                    received_time = str(email.receivedDateTime)
            except:
                received_time = "Unknown"

            # Combine all information
            email_content = f"""
Subject: {subject}

From: {sender_info}
To: {recipient_info}
Received: {received_time}

Body:
{body}
"""
            
            return email_content.strip()
            
        except Exception as e:
            logger.error(f"Error preparing email content: {str(e)}")
            return ""

    def _get_sender_info(self, email) -> str:
        """Extract sender information from email."""
        try:
            sender = getattr(email, 'sender', None)
            if sender:
                if hasattr(sender, 'address') and hasattr(sender, 'name'):
                    return f"{sender.name} <{sender.address}>"
                elif hasattr(sender, 'address'):
                    return sender.address
                else:
                    return str(sender)
            
            # Fallback: try 'from' attribute
            from_field = getattr(email, 'from', None)
            if from_field:
                return str(from_field)
                
            return "Unknown sender"
        except Exception as e:
            logger.warning(f"Error extracting sender info: {str(e)}")
            return "Unknown sender"

    def _get_recipient_info(self, email) -> str:
        """Extract recipient information from email."""
        try:
            recipients = []
            
            # Try to get 'to' recipients
            if hasattr(email, 'to'):
                to_recipients = email.to
                if isinstance(to_recipients, list):
                    for recipient in to_recipients:
                        if hasattr(recipient, 'address'):
                            recipients.append(recipient.address)
                        else:
                            recipients.append(str(recipient))
                elif to_recipients:
                    recipients.append(str(to_recipients))
            
            return ", ".join(recipients) if recipients else "Unknown recipients"
        except Exception as e:
            logger.warning(f"Error extracting recipient info: {str(e)}")
            return "Unknown recipients"

    async def _mark_email_as_read(self, email):
        """Mark an email as read."""
        try:
            # Try the direct method first
            if hasattr(email, 'mark_as_read'):
                email.mark_as_read()
                return
            
            # Try setting isRead property
            if hasattr(email, 'is_read'):
                email.is_read = True
                if hasattr(email, 'save'):
                    email.save()
                return
            
            # Fallback to API call
            email_id = getattr(email, 'object_id', getattr(email, 'id', None))
            if email_id:
                await self._mark_email_read_api(email_id)
                
        except Exception as e:
            logger.warning(f"Could not mark email as read: {str(e)}")

    async def _mark_email_read_api(self, email_id: str):
        """Mark email as read using direct API call."""
        try:
            url = self.account.protocol.service_url
            if self.user_id:
                update_endpoint = f'{url}/users/{self.user_id}/messages/{email_id}'
            else:
                update_endpoint = f'{url}/me/messages/{email_id}'
            
            data = {"isRead": True}
            response = self.account.con.patch(update_endpoint, json=data)
            response.raise_for_status()
            logger.debug(f"Marked email {email_id} as read via API")
        except Exception as e:
            logger.error(f"Failed to mark email {email_id} as read via API: {str(e)}")

    async def process_batch(self, batch_size: int = None, user_id: str = None) -> list:
        """Process a batch of emails for a specific user."""
        try:
            target_user_id = user_id or self.user_id
            batch_size = batch_size or settings.BATCH_SIZE_PER_ACCOUNT
            
            logger.info(f"Starting batch processing for user {target_user_id} (batch size: {batch_size})")
            
            # Fetch unread emails
            emails = await self.fetch_unread_emails(batch_size, target_user_id)
            
            if not emails:
                logger.info(f"No unread emails found for user {target_user_id}")
                return []
            
            logger.info(f"Processing {len(emails)} emails for user {target_user_id}")
            
            # Process emails concurrently
            tasks = [self.process_email(email, target_user_id) for email in emails]
            results = await asyncio.gather(*tasks, return_exceptions=True)
            
            # Filter out None results and exceptions
            valid_results = []
            for result in results:
                if isinstance(result, Exception):
                    logger.error(f"Error in batch processing: {str(result)}")
                elif result is not None:
                    valid_results.append(result)
            
            logger.info(f"Successfully processed {len(valid_results)} emails for user {target_user_id}")
            return valid_results
            
        except Exception as e:
            logger.error(f"Error in batch processing for user {target_user_id}: {str(e)}")
            return []


class MultiAccountEmailProcessor:
    """Processor for handling multiple email accounts simultaneously."""
    
    def __init__(self):
        self.email_accounts = settings.get_email_accounts()
        self.processors = {}
        logger.info(f"Initialized multi-account processor for {len(self.email_accounts)} accounts: {self.email_accounts}")
    
    def get_processor(self, user_id: str) -> EmailProcessor:
        """Get or create an EmailProcessor for a specific user."""
        if user_id not in self.processors:
            self.processors[user_id] = EmailProcessor(user_id)
        return self.processors[user_id]
    
    async def process_account(self, user_id: str, batch_size: int = None) -> Dict:
        """Process emails for a single account."""
        try:
            processor = self.get_processor(user_id)
            results = await processor.process_batch(batch_size, user_id)
            
            return {
                "user_id": user_id,
                "success": True,
                "processed_count": len(results),
                "results": results,
                "error": None
            }
        except Exception as e:
            logger.error(f"Error processing account {user_id}: {str(e)}")
            return {
                "user_id": user_id,
                "success": False,
                "processed_count": 0,
                "results": [],
                "error": str(e)
            }
    
    async def process_all_accounts(self, batch_size_per_account: int = None) -> List[Dict]:
        """Process emails for all configured accounts concurrently."""
        try:
            batch_size = batch_size_per_account or settings.BATCH_SIZE_PER_ACCOUNT
            max_concurrent = settings.MAX_CONCURRENT_ACCOUNTS
            
            logger.info(f"Processing {len(self.email_accounts)} accounts with max {max_concurrent} concurrent")
            
            # Create semaphore to limit concurrent processing
            semaphore = asyncio.Semaphore(max_concurrent)
            
            async def process_with_semaphore(user_id: str):
                async with semaphore:
                    return await self.process_account(user_id, batch_size)
            
            # Process accounts concurrently
            tasks = [process_with_semaphore(user_id) for user_id in self.email_accounts]
            results = await asyncio.gather(*tasks, return_exceptions=True)
            
            # Handle results and exceptions
            final_results = []
            for i, result in enumerate(results):
                if isinstance(result, Exception):
                    logger.error(f"Exception processing account {self.email_accounts[i]}: {str(result)}")
                    final_results.append({
                        "user_id": self.email_accounts[i],
                        "success": False,
                        "processed_count": 0,
                        "results": [],
                        "error": str(result)
                    })
                else:
                    final_results.append(result)
            
            # Log summary
            total_processed = sum(r["processed_count"] for r in final_results)
            successful_accounts = sum(1 for r in final_results if r["success"])
            
            logger.info(f"Multi-account processing complete: {successful_accounts}/{len(self.email_accounts)} accounts successful, {total_processed} total emails processed")
            
            return final_results
            
        except Exception as e:
            logger.error(f"Error in multi-account processing: {str(e)}")
            raise