from flask import Flask, request, jsonify
import os
import base64
import re
from datetime import datetime
import google.generativeai as genai
from dotenv import load_dotenv
import logging
import io
from PIL import Image
import fitz  # PyMuPDF for PDF processing
from docx import Document
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import tempfile
import json

logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

load_dotenv()

api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    raise ValueError("GEMINI_API_KEY not found in environment variables")

genai.configure(api_key=api_key)
model = genai.GenerativeModel('gemini-1.5-flash')

app = Flask(__name__)

def enhance_image_for_ocr(image_data):
    """Enhance image quality for better OCR results"""
    try:
        # Convert base64 to PIL Image
        if isinstance(image_data, str):
            image_bytes = base64.b64decode(image_data)
        else:
            image_bytes = image_data
            
        img = Image.open(io.BytesIO(image_bytes))
        
        # Convert to RGB if needed
        if img.mode in ('RGBA', 'LA'):
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'RGBA':
                background.paste(img, mask=img.split()[-1])
            else:
                background.paste(img)
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')
        
        # Enhance image quality
        width, height = img.size
        
        # Resize if image is too small (improve OCR accuracy)
        if width < 800 or height < 600:
            scale_factor = max(800/width, 600/height)
            new_width = int(width * scale_factor)
            new_height = int(height * scale_factor)
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # Save enhanced image
        output_buffer = io.BytesIO()
        img.save(output_buffer, format='PNG', quality=95)
        enhanced_data = output_buffer.getvalue()
        
        return base64.b64encode(enhanced_data).decode('utf-8')
    except Exception as e:
        logger.warning(f"Image enhancement failed, using original: {str(e)}")
        return image_data if isinstance(image_data, str) else base64.b64encode(image_data).decode('utf-8')

def encode_image_to_base64(image_data):
    """Encode image data to base64 string with enhancement"""
    if isinstance(image_data, bytes):
        enhanced_data = enhance_image_for_ocr(image_data)
        return enhanced_data
    else:
        enhanced_data = enhance_image_for_ocr(image_data.read())
        return enhanced_data

def extract_text_from_word(file_stream):
    """Extract text from Word document with better formatting"""
    try:
        doc = Document(file_stream)
        text_content = []
        
        # Extract paragraphs with better formatting
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                # Preserve some formatting cues
                text = paragraph.text.strip()
                if paragraph.style.name.startswith('Heading'):
                    text = f"[HEADER] {text}"
                text_content.append(text)
        
        # Extract tables with improved structure
        for table in doc.tables:
            table_content = ["[TABLE START]"]
            for row_idx, row in enumerate(table.rows):
                row_text = []
                for cell in row.cells:
                    cell_text = cell.text.strip().replace('\n', ' ')
                    if cell_text:
                        row_text.append(cell_text)
                    else:
                        row_text.append("")
                if row_text and any(cell for cell in row_text):
                    if row_idx == 0:  # Likely header row
                        table_content.append(f"[HEADER ROW] {' | '.join(row_text)}")
                    else:
                        table_content.append(f"[DATA ROW] {' | '.join(row_text)}")
            table_content.append("[TABLE END]")
            text_content.extend(table_content)
        
        return "\n".join(text_content)
    except Exception as e:
        logger.error(f"Error extracting text from Word document: {str(e)}")
        raise Exception(f"Failed to extract text from Word document: {str(e)}")

def extract_data_from_excel_with_formulas(file_stream):
    """Extract data from Excel file with improved structure recognition"""
    try:
        workbook_formulas = openpyxl.load_workbook(file_stream, data_only=False)
        file_stream.seek(0)
        workbook_values = openpyxl.load_workbook(file_stream, data_only=True)
        
        text_content = []
        
        for sheet_name in workbook_formulas.sheetnames:
            sheet_formulas = workbook_formulas[sheet_name]
            sheet_values = workbook_values[sheet_name]
            sheet_content = [f"[SHEET: {sheet_name}]"]
            
            max_row = sheet_formulas.max_row
            max_col = sheet_formulas.max_column
            
            # Try to identify header rows and data structure
            header_detected = False
            
            for row in range(1, max_row + 1):
                row_data = []
                has_content = False
                
                for col in range(1, max_col + 1):
                    cell_formula = sheet_formulas.cell(row, col)
                    cell_value = sheet_values.cell(row, col)
                    
                    cell_content = ""
                    if cell_formula.value is not None:
                        if isinstance(cell_formula.value, str) and cell_formula.value.startswith('='):
                            calculated_value = cell_value.value if cell_value.value is not None else 0
                            cell_content = f"{cell_formula.value}[={calculated_value}]"
                        else:
                            cell_content = str(cell_formula.value)
                        has_content = True
                    elif cell_value.value is not None:
                        cell_content = str(cell_value.value)
                        has_content = True
                    
                    row_data.append(cell_content)
                
                if has_content:
                    # Detect if this might be a header row
                    if not header_detected and row <= 3:  # Check first 3 rows for headers
                        text_values = [cell for cell in row_data if cell and not cell.replace('.', '').replace('-', '').isdigit()]
                        if len(text_values) >= len([cell for cell in row_data if cell]) * 0.7:  # 70% text content
                            sheet_content.append(f"[HEADER ROW] {' | '.join(row_data)}")
                            header_detected = True
                            continue
                    
                    sheet_content.append(f"[DATA ROW] {' | '.join(row_data)}")
            
            if len(sheet_content) > 1:
                text_content.extend(sheet_content)
        
        return "\n".join(text_content)
    except Exception as e:
        logger.error(f"Error extracting data from Excel file: {str(e)}")
        raise Exception(f"Failed to extract data from Excel file: {str(e)}")

def extract_images_from_pdf(file_stream):
    """Extract images from PDF with better quality"""
    try:
        pdf_document = fitz.open(stream=file_stream.read(), filetype="pdf")
        images = []
        
        for page_num in range(pdf_document.page_count):
            page = pdf_document[page_num]
            image_list = page.get_images()
            
            for img_index, img in enumerate(image_list):
                xref = img[0]
                pix = fitz.Pixmap(pdf_document, xref)
                
                # Skip very small images (likely decorative)
                if pix.width < 100 or pix.height < 100:
                    continue
                
                if pix.n - pix.alpha < 4:  # GRAY or RGB
                    img_data = pix.tobytes("png")
                    images.append(img_data)
                else:  # CMYK: convert to RGB first
                    pix1 = fitz.Pixmap(fitz.csRGB, pix)
                    img_data = pix1.tobytes("png")
                    images.append(img_data)
                    pix1 = None
                
                pix = None
        
        pdf_document.close()
        return images
    except Exception as e:
        logger.error(f"Error extracting images from PDF: {str(e)}")
        raise Exception(f"Failed to extract images from PDF: {str(e)}")

def extract_text_from_pdf(file_stream):
    """Extract text from PDF with better structure preservation"""
    try:
        pdf_document = fitz.open(stream=file_stream.read(), filetype="pdf")
        text_content = []
        
        for page_num in range(pdf_document.page_count):
            page = pdf_document[page_num]
            
            # Get text with layout information
            blocks = page.get_text("dict")
            page_text = [f"[PAGE {page_num + 1}]"]
            
            for block in blocks["blocks"]:
                if "lines" in block:
                    block_text = []
                    for line in block["lines"]:
                        line_text = ""
                        for span in line["spans"]:
                            text = span["text"].strip()
                            if text:
                                # Detect if text might be a header (larger font)
                                if span["size"] > 12:
                                    text = f"[LARGE TEXT] {text}"
                                line_text += text + " "
                        if line_text.strip():
                            block_text.append(line_text.strip())
                    
                    if block_text:
                        page_text.extend(block_text)
            
            if len(page_text) > 1:
                text_content.extend(page_text)
        
        pdf_document.close()
        return "\n".join(text_content)
    except Exception as e:
        logger.error(f"Error extracting text from PDF: {str(e)}")
        raise Exception(f"Failed to extract text from PDF: {str(e)}")

def process_file_content(file_stream, filename):
    """Process different file types and extract content with enhancements"""
    file_extension = filename.rsplit('.', 1)[1].lower()
    
    try:
        if file_extension in ['png', 'jpg', 'jpeg', 'webp']:
            image_data = encode_image_to_base64(file_stream)
            return {"type": "image", "data": image_data, "extension": file_extension}
        
        elif file_extension == 'pdf':
            file_stream.seek(0)
            pdf_images = extract_images_from_pdf(file_stream)
            
            if pdf_images:
                # Use the largest/best quality image
                best_image = max(pdf_images, key=len)  # Largest file size often means best quality
                image_data = encode_image_to_base64(best_image)
                return {"type": "image", "data": image_data, "extension": file_extension}
            else:
                file_stream.seek(0)
                text_content = extract_text_from_pdf(file_stream)
                return {"type": "text", "data": text_content, "extension": file_extension}
        
        elif file_extension in ['doc', 'docx']:
            text_content = extract_text_from_word(file_stream)
            return {"type": "text", "data": text_content, "extension": file_extension}
        
        elif file_extension in ['xls', 'xlsx']:
            text_content = extract_data_from_excel_with_formulas(file_stream)
            return {"type": "text", "data": text_content, "extension": file_extension}
        
        else:
            raise Exception("Unsupported file type")
            
    except Exception as e:
        raise Exception(f"Error processing file: {str(e)}")

def validate_and_clean_extracted_data(extracted_data):
    """Validate and clean extracted data for better quality"""
    cleaned_data = extracted_data.copy()
    
    # Clean and validate date
    if cleaned_data.get('date'):
        date_str = str(cleaned_data['date']).strip()
        # Try to parse and reformat date for consistency
        try:
            # Handle various date formats
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%m-%d-%Y', '%d-%m-%Y']:
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    cleaned_data['date'] = parsed_date.strftime('%Y-%m-%d')
                    break
                except ValueError:
                    continue
        except:
            pass  # Keep original if parsing fails
    
    # Clean numeric fields
    numeric_fields = ['subtotal', 'tax', 'total']
    for field in numeric_fields:
        if cleaned_data.get(field):
            value = str(cleaned_data[field])
            # Remove currency symbols and clean numeric values
            cleaned_value = re.sub(r'[^\d.,\-]', '', value)
            cleaned_value = cleaned_value.replace(',', '')
            try:
                float(cleaned_value)
                cleaned_data[field] = cleaned_value
            except ValueError:
                pass  # Keep original if not numeric
    
    # Clean and validate items
    if cleaned_data.get('items') and isinstance(cleaned_data['items'], list):
        cleaned_items = []
        for item in cleaned_data['items']:
            if isinstance(item, dict):
                cleaned_item = {}
                for key, value in item.items():
                    if value and str(value).strip() and str(value).strip().lower() not in ['null', 'none', '']:
                        if key in ['quantity', 'unit_price', 'amount']:
                            # Clean numeric values in items
                            cleaned_value = re.sub(r'[^\d.,\-]', '', str(value))
                            cleaned_value = cleaned_value.replace(',', '')
                            try:
                                float(cleaned_value)
                                cleaned_item[key] = cleaned_value
                            except ValueError:
                                cleaned_item[key] = str(value).strip()
                        else:
                            cleaned_item[key] = str(value).strip()
                
                # Only add item if it has meaningful content
                if cleaned_item.get('description') or cleaned_item.get('quantity'):
                    cleaned_items.append(cleaned_item)
        
        cleaned_data['items'] = cleaned_items
    
    return cleaned_data

def get_extraction_prompt(file_extension):
    """Get enhanced extraction prompt based on file type"""
    base_prompt = """
    You are an expert invoice data extraction system. Extract the following information with MAXIMUM ACCURACY:

    REQUIRED FIELDS:
    1. Invoice Number (look for: Invoice #, Invoice No, Bill #, etc.)
    2. Date (look for: Date, Invoice Date, Bill Date, etc.)
    3. Customer Name (look for: Bill To, Customer, Client, etc.)
    4. Customer Contact (phone/email)
    5. Customer Address (complete address)
    6. Items with details:
       - Quantity (Qty, Quantity, Units)
       - Description (Item, Product, Service, Description)
       - Unit Price (Price, Rate, Unit Cost, Each)
       - Amount (Total, Line Total, Subtotal per item)
    7. Subtotal (before tax)
    8. Tax amount (Tax, VAT, GST, etc.)
    9. Total amount (Grand Total, Final Amount, etc.)

    EXTRACTION RULES:
    - Extract EXACT values as they appear in the document
    - Look for currency symbols and preserve them in context
    - For tables, identify header rows vs data rows
    - Pay attention to formatting cues like bold text, larger fonts
    - If multiple similar fields exist, choose the most prominent one
    - For calculations, use the final computed values shown
    - Preserve decimal precision as shown
    - Handle different number formats (1,234.56 vs 1234.56)
    """
    
    if file_extension in ['xls', 'xlsx']:
        specific_instructions = """
        EXCEL-SPECIFIC INSTRUCTIONS:
        - Extract values from cells, not formulas
        - Look for [HEADER ROW] and [DATA ROW] markers
        - For formula cells showing "=formula[=result]", use the result value
        - Pay attention to sheet structure and table layout
        - Handle merged cells appropriately
        - Identify the main invoice data vs other sheets
        """
    else:
        specific_instructions = """
        DOCUMENT-SPECIFIC INSTRUCTIONS:
        - For images: Focus on text regions, handle various fonts and sizes
        - For PDFs: Use structural markers like [LARGE TEXT] for headers
        - For Word docs: Use [HEADER] and [TABLE] markers for structure
        - Look for visual cues like borders, spacing, alignment
        - Handle multi-column layouts appropriately
        """
    
    json_format = """
    FORMAT RESPONSE AS VALID JSON:
    {
        "invoice_number": "extracted_value_or_null",
        "date": "extracted_value_or_null",
        "customer_name": "extracted_value_or_null",
        "customer_contact": "extracted_value_or_null",
        "customer_address": "extracted_value_or_null",
        "items": [
            {
                "quantity": "extracted_value_or_null",
                "description": "extracted_value_or_null",
                "unit_price": "extracted_value_or_null",
                "amount": "extracted_value_or_null"
            }
        ],
        "subtotal": "extracted_value_or_null",
        "tax": "extracted_value_or_null",
        "total": "extracted_value_or_null"
    }

    IMPORTANT: 
    - Return ONLY the JSON, no other text
    - Use null for missing fields
    - Ensure all numbers are strings to preserve formatting
    - Include all line items found in the items array
    - Double-check calculations make sense
    """
    
    return base_prompt + specific_instructions + json_format

def calculate_item_amounts_excel_only(items, file_extension):
    """Calculate item amounts only for Excel files with better validation"""
    if file_extension not in ['xls', 'xlsx']:
        return items
    
    calculated_items = []
    
    for item in items:
        calculated_item = item.copy()
        
        try:
            quantity_str = str(item.get('quantity', '0')).replace(',', '')
            unit_price_str = str(item.get('unit_price', '0')).replace(',', '')
            
            # Extract numeric values more carefully
            quantity = 0
            unit_price = 0
            
            # Handle formula results
            if '[=' in quantity_str:
                match = re.search(r'\[=([^\]]+)\]', quantity_str)
                if match:
                    quantity = float(match.group(1))
                else:
                    quantity = float(re.sub(r'[^\d.-]', '', quantity_str))
            else:
                quantity = float(re.sub(r'[^\d.-]', '', quantity_str)) if quantity_str else 0
            
            if '[=' in unit_price_str:
                match = re.search(r'\[=([^\]]+)\]', unit_price_str)
                if match:
                    unit_price = float(match.group(1))
                else:
                    unit_price = float(re.sub(r'[^\d.-]', '', unit_price_str))
            else:
                unit_price = float(re.sub(r'[^\d.-]', '', unit_price_str)) if unit_price_str else 0
            
            # Calculate amount if missing or formula-based
            amount = item.get('amount', '')
            if not amount or '=' in str(amount) or '[=' in str(amount):
                calculated_amount = round(quantity * unit_price, 2)
                calculated_item['amount'] = str(calculated_amount)
            else:
                # Clean existing amount
                amount_str = str(amount).replace(',', '')
                if '[=' in amount_str:
                    match = re.search(r'\[=([^\]]+)\]', amount_str)
                    if match:
                        calculated_item['amount'] = str(float(match.group(1)))
                else:
                    try:
                        calculated_item['amount'] = str(float(re.sub(r'[^\d.-]', '', amount_str)))
                    except:
                        calculated_item['amount'] = str(amount)
                        
        except (ValueError, TypeError) as e:
            logger.warning(f"Error calculating item amount: {e}")
            pass
            
        calculated_items.append(calculated_item)
    
    return calculated_items

def calculate_totals_excel_only(items, extracted_data, file_extension):
    """Calculate totals only for Excel files with better validation"""
    if file_extension not in ['xls', 'xlsx']:
        return {
            'subtotal': extracted_data.get('subtotal'),
            'tax': extracted_data.get('tax'),
            'total': extracted_data.get('total')
        }
    
    try:
        # Calculate subtotal from items
        subtotal = 0
        for item in items:
            try:
                amount_str = str(item.get('amount', '0')).replace(',', '')
                amount = float(re.sub(r'[^\d.-]', '', amount_str))
                subtotal += amount
            except (ValueError, TypeError):
                pass
        
        # Handle tax with formula support
        tax_value = 0
        tax_field = extracted_data.get('tax')
        if tax_field and str(tax_field).strip() not in ['null', 'None', '']:
            tax_str = str(tax_field)
            try:
                if '[=' in tax_str:
                    match = re.search(r'\[=([^\]]+)\]', tax_str)
                    if match:
                        tax_value = float(match.group(1))
                else:
                    tax_value = float(re.sub(r'[^\d.-]', '', tax_str))
            except (ValueError, TypeError):
                pass
        
        # Calculate total
        total = round(subtotal + tax_value, 2)
        
        return {
            'subtotal': str(round(subtotal, 2)),
            'tax': str(round(tax_value, 2)) if tax_value > 0 else extracted_data.get('tax'),
            'total': str(total)
        }
        
    except Exception as e:
        logger.error(f"Error calculating totals: {str(e)}")
        return {
            'subtotal': extracted_data.get('subtotal'),
            'tax': extracted_data.get('tax'),
            'total': extracted_data.get('total')
        }

def extract_invoice_data_with_gemini(content_data):
    """Use Google Gemini AI to extract invoice data with enhanced prompting"""
    try:
        file_extension = content_data.get('extension', 'unknown')
        prompt = get_extraction_prompt(file_extension)
        
        if content_data["type"] == "image":
            response = model.generate_content([prompt, {"mime_type": "image/png", "data": content_data["data"]}])
        elif content_data["type"] == "text":
            full_prompt = f"{prompt}\n\nDocument content:\n{content_data['data']}"
            response = model.generate_content(full_prompt)
        else:
            raise Exception("Invalid content type")
        
        response_text = response.text
        logger.info(f"Raw AI response: {response_text[:500]}...")  # Log first 500 chars
        
        # Extract JSON from response with better parsing
        json_match = re.search(r'({[\s\S]*})', response_text)
        if json_match:
            json_str = json_match.group(1)
            json_str = json_str.replace('```json', '').replace('```', '').strip()
            
            try:
                extracted_data = json.loads(json_str)
                
                # Validate and clean the data
                extracted_data = validate_and_clean_extracted_data(extracted_data)
                
                # Post-process ONLY for Excel files
                if 'items' in extracted_data and extracted_data['items']:
                    calculated_items = calculate_item_amounts_excel_only(extracted_data['items'], file_extension)
                    extracted_data['items'] = calculated_items
                    
                    calculated_totals = calculate_totals_excel_only(calculated_items, extracted_data, file_extension)
                    extracted_data.update(calculated_totals)
                
                extracted_data['file_extension'] = file_extension
                
                return extracted_data
            except json.JSONDecodeError as e:
                logger.error(f"JSON parsing error: {e}")
                logger.error(f"Problematic JSON: {json_str}")
                raise Exception("Failed to parse extracted data as JSON")
        else:
            logger.error(f"No JSON found in response: {response_text}")
            raise Exception("No structured data found in the AI response")
    
    except Exception as e:
        logger.error(f"Error in Gemini API: {str(e)}")
        raise Exception(f"AI processing error: {str(e)}")

def create_api_response(status, message, data=None):
    """Create standardized API response"""
    response = {
        "status": status,
        "message": message,
        "data": data
    }
    return response

@app.route('/extract-invoice', methods=['POST'])
def extract_invoice():
    """API endpoint to extract data from invoice files with enhanced processing"""
    
    try:
        if 'file' not in request.files:
            response = create_api_response("error", "No file provided", None)
            return jsonify(response), 400
        
        uploaded_file = request.files['file']
        
        if uploaded_file.filename == '':
            response = create_api_response("error", "Empty filename", None)
            return jsonify(response), 400

        allowed_extensions = {'png', 'jpg', 'jpeg', 'webp', 'pdf', 'doc', 'docx', 'xls', 'xlsx'}
        file_extension = uploaded_file.filename.rsplit('.', 1)[1].lower() if '.' in uploaded_file.filename else ''
        
        if not (file_extension in allowed_extensions):
            response = create_api_response(
                "error", 
                "File type not allowed. Please upload PNG, JPG, JPEG, WEBP, PDF, DOC, DOCX, XLS, or XLSX files", 
                None
            )
            return jsonify(response), 400

        file_stream = io.BytesIO(uploaded_file.read())

        try:
            content_data = process_file_content(file_stream, uploaded_file.filename)
            logger.info(f"Successfully processed {file_extension} file")
        except Exception as e:
            logger.error(f"File processing error: {str(e)}")
            response = create_api_response("error", str(e), None)
            return jsonify(response), 400

        try:
            extracted_data = extract_invoice_data_with_gemini(content_data)

            logger.info(f"Extraction completed for {file_extension} file")
            logger.info(f"Items extracted: {len(extracted_data.get('items', []))}")
            

            if file_extension in ['xls', 'xlsx']:
                message = f"Invoice data extracted successfully from {file_extension.upper()} file with enhanced calculation processing"
            else:
                message = f"Invoice data extracted successfully from {file_extension.upper()} file with enhanced OCR and structure recognition"
            
            response = create_api_response("success", message, extracted_data)
            return jsonify(response), 200
            
        except Exception as e:
            logger.error(f"Extraction error: {str(e)}")
            response = create_api_response("error", str(e), None)
            return jsonify(response), 500
    
    except Exception as e:
        logger.error(f"Unexpected error processing invoice: {str(e)}")
        response = create_api_response("error", f"Unexpected error: {str(e)}", None)
        return jsonify(response), 500

@app.errorhandler(404)
def not_found(error):
    """Handle 404 errors"""
    response = create_api_response("error", "Endpoint not found", None)
    return jsonify(response), 404

@app.errorhandler(405)
def method_not_allowed(error):
    """Handle 405 errors"""
    response = create_api_response("error", "Method not allowed", None)
    return jsonify(response), 405

@app.errorhandler(500)
def internal_error(error):
    """Handle 500 errors"""
    response = create_api_response("error", "Internal server error", None)
    return jsonify(response), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)