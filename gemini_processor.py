import google.generativeai as genai
import json
import re
from tenacity import retry, stop_after_attempt, wait_exponential
from loguru import logger
from config import settings
import time
import asyncio
from datetime import datetime

class RateLimiter:
    """Token bucket rate limiter for managing API request rates."""
    
    def __init__(self, rpm_limit: int = 30, tpm_limit: int = 1_000_000):
        self.rpm_limit = rpm_limit
        self.tpm_limit = tpm_limit
        self.request_tokens = rpm_limit  # Current available request tokens
        self.total_tokens = tpm_limit    # Current available total tokens
        self.last_refill = time.time()
        self.lock = asyncio.Lock()
    
    async def acquire(self, tokens_needed: int = 1) -> bool:
        """
        Attempt to acquire tokens for a request.
        
        Args:
            tokens_needed: Number of tokens needed for this request
            
        Returns:
            bool: True if tokens were acquired, False if should retry later
        """
        async with self.lock:
            now = time.time()
            time_passed = now - self.last_refill
            
            # Refill tokens based on time passed (up to max)
            if time_passed >= 1:  # Refill every second
                # Calculate tokens to add (proportional to time passed, up to 1 minute)
                minutes_passed = min(1, time_passed / 60)
                request_tokens_to_add = int(self.rpm_limit * minutes_passed)
                total_tokens_to_add = int(self.tpm_limit * minutes_passed)
                
                self.request_tokens = min(self.rpm_limit, self.request_tokens + request_tokens_to_add)
                self.total_tokens = min(self.tpm_limit, self.total_tokens + total_tokens_to_add)
                self.last_refill = now
            
            # Check if we have enough tokens
            if self.request_tokens >= 1 and self.total_tokens >= tokens_needed:
                self.request_tokens -= 1
                self.total_tokens -= tokens_needed
                return True
            
            return False
    
    async def wait_for_tokens(self, tokens_needed: int = 1):
        """Wait until tokens are available."""
        while not await self.acquire(tokens_needed):
            await asyncio.sleep(1)  # Wait 1 second before retrying

class GeminiProcessor:
    def __init__(self):
        self.model = self._initialize_model()
        self.prompt_template = self._get_prompt_template()
        self.rate_limiter = RateLimiter()
        
    def _initialize_model(self):
        """Initialize the Gemini model with proper configuration."""
        try:
            genai.configure(api_key=settings.GEMINI_API_KEY)
            return genai.GenerativeModel('gemini-1.5-flash')  # Use flash model for better reliability
        except Exception as e:
            logger.error(f"Failed to initialize Gemini model: {str(e)}")
            raise
            
    def _get_prompt_template(self):
        """Return the standardized prompt template for extraction."""
        return """Extract business information from the email and return ONLY valid JSON.

Return ONLY this JSON structure (no explanations, no markdown):
{
  "customer_name": null,
  "company_name": null,
  "email": null,
  "contact_number": null,
  "shipping_from_address": null,
  "shipping_to_address": null,
  "shipping_from_city": null,
  "shipping_to_city": null,
  "shipping_from_country": null,
  "shipping_to_country": null,
  "country": null,
  "currency": null,
  "products": [],
  "quantity": null,
  "mode_of_transportation": null,
  "order_date": null,
  "shipped_date": null,
  "expected_delivery_date": null,
  "actual_delivery_date": null,
  "email_received_date": null,
  "confidence_score": 0.0
}

Fill in actual values where found, otherwise keep as null.
For dates, use ISO format (YYYY-MM-DD) when possible.
For shipping addresses, extract full addresses including street, city, postal code.
The "email" field should contain the sender's/client's email address (who sent this email).

EMAIL CONTENT:
{email_content}"""
    
    def _clean_response_text(self, text: str) -> str:
        """Clean and fix the response text to extract valid JSON."""
        try:
            if not text:
                logger.warning("Empty response from Gemini")
                return json.dumps(self._get_default_structure())

            # Log the raw response for debugging
            logger.info(f"Raw Gemini response ({len(text)} chars): {repr(text[:200])}...")
            
            # Handle various response formats
            original_text = text
            
            # Remove markdown code blocks
            text = re.sub(r'```json\s*', '', text, flags=re.IGNORECASE)
            text = re.sub(r'```\s*$', '', text, flags=re.MULTILINE)
            text = re.sub(r'^```json\s*', '', text, flags=re.MULTILINE)
            text = text.strip()
            
            # Look for JSON object boundaries
            start_pos = -1
            end_pos = -1
            
            # Find the first opening brace
            for i, char in enumerate(text):
                if char == '{':
                    start_pos = i
                    break
            
            if start_pos == -1:
                logger.warning(f"No opening brace found in response: {repr(text[:100])}")
                # If no JSON found, check if it's a partial response
                if any(field in text.lower() for field in ['customer_name', 'company_name', 'email']):
                    logger.warning("Detected partial field names - response was likely truncated")
                return json.dumps(self._get_default_structure())
            
            # Find the matching closing brace
            brace_count = 0
            for i in range(start_pos, len(text)):
                if text[i] == '{':
                    brace_count += 1
                elif text[i] == '}':
                    brace_count -= 1
                    if brace_count == 0:
                        end_pos = i
                        break
            
            if end_pos == -1:
                logger.warning(f"No closing brace found - truncated response detected")
                # Try to reconstruct a complete JSON from partial response
                return self._reconstruct_partial_json(text[start_pos:])
            
            # Extract the JSON portion
            json_text = text[start_pos:end_pos + 1]
            logger.debug(f"Extracted JSON candidate: {json_text[:200]}...")
            
            # Try to parse the extracted JSON
            try:
                parsed = json.loads(json_text)
                
                if not isinstance(parsed, dict):
                    logger.warning(f"Response is not a JSON object: {type(parsed)}")
                    return json.dumps(self._get_default_structure())
                
                # Merge with default structure
                result = self._merge_with_default(parsed)
                logger.info(f"Successfully parsed JSON with {sum(1 for v in result.values() if v not in [None, [], ''])} populated fields")
                return json.dumps(result, indent=2)
                
            except json.JSONDecodeError as e:
                logger.warning(f"JSON parsing failed: {str(e)}")
                logger.warning(f"Attempting to fix malformed JSON: {json_text[:100]}...")
                
                # Try to fix and reparse
                fixed_json = self._fix_malformed_json(json_text)
                if fixed_json:
                    return fixed_json
                
                logger.error(f"Could not fix JSON. Original text: {repr(original_text[:200])}")
                return json.dumps(self._get_default_structure())
                
        except Exception as e:
            logger.error(f"Error cleaning response text: {str(e)}")
            return json.dumps(self._get_default_structure())
    
    def _reconstruct_partial_json(self, partial_text: str) -> str:
        """Reconstruct a complete JSON from a partial response."""
        try:
            logger.info(f"Attempting to reconstruct partial JSON from: {repr(partial_text[:100])}")
            
            # Start with default structure
            result = self._get_default_structure()
            
            # Try to extract any field:value pairs from the partial text
            field_patterns = [
                r'"customer_name"\s*:\s*"([^"]*)"',
                r'"company_name"\s*:\s*"([^"]*)"',
                r'"email"\s*:\s*"([^"]*)"',
                r'"contact_number"\s*:\s*"([^"]*)"',
                r'"shipping_from_address"\s*:\s*"([^"]*)"',
                r'"shipping_to_address"\s*:\s*"([^"]*)"',
                r'"shipping_from_city"\s*:\s*"([^"]*)"',
                r'"shipping_to_city"\s*:\s*"([^"]*)"',
                r'"shipping_from_country"\s*:\s*"([^"]*)"',
                r'"shipping_to_country"\s*:\s*"([^"]*)"',
                r'"country"\s*:\s*"([^"]*)"',
                r'"currency"\s*:\s*"([^"]*)"',
                r'"mode_of_transportation"\s*:\s*"([^"]*)"',
                r'"order_date"\s*:\s*"([^"]*)"',
                r'"shipped_date"\s*:\s*"([^"]*)"',
                r'"expected_delivery_date"\s*:\s*"([^"]*)"',
                r'"actual_delivery_date"\s*:\s*"([^"]*)"',
                r'"email_received_date"\s*:\s*"([^"]*)"'
            ]
            
            found_fields = 0
            for pattern in field_patterns:
                match = re.search(pattern, partial_text)
                if match:
                    field_name = pattern.split('"')[1]
                    field_value = match.group(1)
                    if field_value:
                        result[field_name] = field_value
                        found_fields += 1
            
            # Handle products array
            products_match = re.search(r'"products"\s*:\s*\[([^\]]*)\]', partial_text)
            if products_match:
                products_text = products_match.group(1)
                products = [p.strip(' "') for p in products_text.split(',') if p.strip()]
                if products:
                    result['products'] = products
                    found_fields += 1
            
            # Handle confidence score
            confidence_match = re.search(r'"confidence_score"\s*:\s*([0-9.]+)', partial_text)
            if confidence_match:
                try:
                    confidence = float(confidence_match.group(1))
                    result['confidence_score'] = max(0.0, min(1.0, confidence))
                    found_fields += 1
                except ValueError:
                    pass
            
            logger.info(f"Reconstructed {found_fields} fields from partial response")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error reconstructing partial JSON: {str(e)}")
            return json.dumps(self._get_default_structure())
    
    def _fix_malformed_json(self, json_text: str) -> str:
        """Try to fix common JSON formatting issues."""
        try:
            # Remove trailing commas
            json_text = re.sub(r',\s*}', '}', json_text)
            json_text = re.sub(r',\s*]', ']', json_text)
            
            # Fix missing quotes around string values
            json_text = re.sub(r':\s*([a-zA-Z][^,}\]]*)\s*([,}])', r': "\1"\2', json_text)
            
            # Ensure all field names are quoted
            json_text = re.sub(r'([a-zA-Z_][a-zA-Z0-9_]*)\s*:', r'"\1":', json_text)
            
            # Try parsing again
            parsed = json.loads(json_text)
            if isinstance(parsed, dict):
                result = self._merge_with_default(parsed)
                logger.info("Successfully fixed malformed JSON")
                return json.dumps(result, indent=2)
                
        except Exception as e:
            logger.warning(f"Could not fix malformed JSON: {str(e)}")
        
        return None
    
    def _merge_with_default(self, parsed_data: dict) -> dict:
        """Merge parsed data with default structure."""
        default_structure = self._get_default_structure()
        
        for key in default_structure:
            if key in parsed_data and parsed_data[key] is not None:
                # Validate and clean the value
                value = parsed_data[key]
                
                if key == 'products' and not isinstance(value, list):
                    if isinstance(value, str) and value:
                        default_structure[key] = [value]
                    else:
                        default_structure[key] = []
                elif key == 'confidence_score':
                    try:
                        score = float(value)
                        default_structure[key] = max(0.0, min(1.0, score))
                    except (TypeError, ValueError):
                        default_structure[key] = 0.0
                elif isinstance(value, str):
                    cleaned_value = value.strip()
                    default_structure[key] = cleaned_value if cleaned_value else None
                else:
                    default_structure[key] = value
        
        return default_structure

    @retry(
        stop=stop_after_attempt(2),  # Reduced retries since the issue is parsing, not API
        wait=wait_exponential(multiplier=1, min=2, max=6),
        reraise=True
    )
    async def process_email(self, email_content: str) -> dict:
        """Process email content using Gemini AI with enhanced error handling."""
        try:
            logger.info("🔍 Starting Gemini processing...")
            
            # Validate email content
            if not email_content or not isinstance(email_content, str):
                logger.warning("Invalid or empty email content provided")
                return self._get_default_structure()

            # Truncate very long emails to avoid token limits
            if len(email_content) > 8000:  # Conservative limit
                email_content = email_content[:8000] + "\n[Content truncated due to length]"
                logger.info("Email content truncated due to length")

            # Prepare the prompt - use safe string replacement to avoid KeyError from braces in email content
            prompt = self.prompt_template.replace("{email_content}", email_content)
            logger.info(f"📝 Prepared prompt length: {len(prompt)} chars")
            
            # Estimate tokens (rough estimate: 4 chars = 1 token)
            estimated_tokens = len(prompt) // 4
            
            # Wait for rate limit tokens
            logger.info("⏳ Waiting for rate limit tokens...")
            await self.rate_limiter.wait_for_tokens(estimated_tokens)
            
            logger.info(f"🚀 Sending request to Gemini: {len(email_content)} chars email, {estimated_tokens} estimated tokens")
            
            # Generate response with conservative parameters
            try:
                response = await self.model.generate_content_async(
                    prompt,
                    generation_config={
                        'temperature': 0.0,  # Deterministic output
                        'candidate_count': 1,
                        'max_output_tokens': 512,  # Conservative limit to avoid truncation
                        'top_p': 0.9,
                        'top_k': 20
                    },
                    safety_settings=[
                        {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                        {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
                        {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                        {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
                    ]
                )
                logger.info("✅ Received response from Gemini API")
                
            except Exception as api_error:
                logger.error(f"❌ Error during Gemini API call: {str(api_error)}")
                logger.error(f"Error type: {type(api_error)}")
                return self._get_default_structure()
            
            if not response:
                logger.warning("❌ No response object from Gemini API")
                return self._get_default_structure()
                
            if not hasattr(response, 'text'):
                logger.warning(f"❌ Response has no text attribute. Response type: {type(response)}")
                logger.warning(f"Response attributes: {dir(response)}")
                return self._get_default_structure()
                
            if not response.text:
                logger.warning("❌ Empty response text from Gemini API")
                return self._get_default_structure()
            
            logger.info(f"📥 Raw response received: {len(response.text)} chars")
            
            # Clean and parse response
            try:
                logger.info("🧹 Cleaning response text...")
                cleaned_text = self._clean_response_text(response.text)
                logger.info(f"✅ Response cleaned successfully: {len(cleaned_text)} chars")
                
            except Exception as clean_error:
                logger.error(f"❌ Error during response cleaning: {str(clean_error)}")
                return self._get_default_structure()
            
            try:
                logger.info("📊 Parsing JSON...")
                result = json.loads(cleaned_text)
                
                # Log success
                non_null_fields = sum(1 for v in result.values() if v not in [None, [], ''])
                logger.info(f"✅ Gemini extraction completed with {non_null_fields} populated fields")
                
                return result
                
            except json.JSONDecodeError as e:
                logger.error(f"❌ Final JSON parsing failed: {str(e)}")
                logger.error(f"Cleaned response: {cleaned_text[:300]}...")
                return self._get_default_structure()
                
        except Exception as e:
            logger.error(f"❌ Unexpected error in process_email: {str(e)}")
            logger.error(f"Error type: {type(e)}")
            import traceback
            logger.error(f"Traceback: {traceback.format_exc()}")
            return self._get_default_structure()

    def _get_default_structure(self) -> dict:
        """Return default structure for failed extractions."""
        return {
            "customer_name": None,
            "company_name": None,
            "email": None,
            "contact_number": None,
            "shipping_from_address": None,
            "shipping_to_address": None,
            "shipping_from_city": None,
            "shipping_to_city": None,
            "shipping_from_country": None,
            "shipping_to_country": None,
            "country": None,
            "currency": None,
            "products": [],
            "quantity": None,
            "mode_of_transportation": None,
            "order_date": None,
            "shipped_date": None,
            "expected_delivery_date": None,
            "actual_delivery_date": None,
            "email_received_date": None,
            "confidence_score": 0.0
        }