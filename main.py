import asyncio
import schedule
import time
from contextlib import asynccontextmanager
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse
from loguru import logger
import uvicorn
from datetime import datetime
import json
import os

from config import settings
from email_processor import MultiAccountEmailProcessor

# Global variables
multi_account_processor = None
processing_stats = {
    "total_processed": 0,
    "successful": 0,
    "failed": 0,
    "last_run": None,
    "average_confidence": 0.0
}

@asynccontextmanager
async def lifespan(app: FastAPI):
    """FastAPI lifespan event handler to replace deprecated startup/shutdown events."""
    global multi_account_processor
    
    # Startup
    logger.info("Starting application...")
    multi_account_processor = MultiAccountEmailProcessor()
    
    # Initialize scheduler
    schedule.every(settings.PROCESSING_INTERVAL).seconds.do(
        lambda: asyncio.create_task(process_emails())
    )
    
    # Start scheduler task
    scheduler_task = asyncio.create_task(run_scheduler())
    
    yield  # Application is running
    
    # Shutdown
    logger.info("Shutting down application...")
    scheduler_task.cancel()
    try:
        await scheduler_task
    except asyncio.CancelledError:
        pass

# Initialize FastAPI app with lifespan
app = FastAPI(
    title="Email Extraction API",
    description="API for automated email data extraction using Gemini AI",
    version="1.0.0",
    lifespan=lifespan
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static files
app.mount("/static", StaticFiles(directory="static"), name="static")

# Create output directory if it doesn't exist
os.makedirs("output", exist_ok=True)

async def process_emails():
    """Main processing function that runs on schedule."""
    global multi_account_processor, processing_stats
    
    try:
        logger.info("Starting scheduled multi-account email processing")
        
        # Process all accounts
        account_results = await multi_account_processor.process_all_accounts()
        
        # Aggregate results from all accounts
        all_results = []
        successful_accounts = 0
        failed_accounts = 0
        
        for account_result in account_results:
            if account_result["success"]:
                successful_accounts += 1
                all_results.extend(account_result["results"])
            else:
                failed_accounts += 1
                logger.error(f"Failed to process account {account_result['user_id']}: {account_result['error']}")
        
        # Update statistics
        processing_stats["total_processed"] += len(all_results)
        processing_stats["successful"] += len(all_results)
        processing_stats["last_run"] = datetime.utcnow().isoformat()
        processing_stats["accounts_processed"] = len(account_results)
        processing_stats["successful_accounts"] = successful_accounts
        processing_stats["failed_accounts"] = failed_accounts
        
        if all_results:
            # Calculate average confidence
            confidences = [r.get("confidence_score", 0) for r in all_results]
            processing_stats["average_confidence"] = sum(confidences) / len(confidences)
            
            # Save results to file
            timestamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
            output_file = f"output/extractions_{timestamp}.json"
            
            # Add summary information
            output_data = {
                "summary": {
                    "timestamp": timestamp,
                    "total_accounts": len(account_results),
                    "successful_accounts": successful_accounts,
                    "failed_accounts": failed_accounts,
                    "total_emails_processed": len(all_results),
                    "average_confidence": processing_stats["average_confidence"]
                },
                "account_results": account_results,
                "all_extractions": all_results
            }
            
            with open(output_file, "w") as f:
                json.dump(output_data, f, indent=2)
                
            logger.info(f"Saved {len(all_results)} extractions from {successful_accounts} accounts to {output_file}")
        else:
            logger.info("No emails processed from any account")
            
    except Exception as e:
        logger.error(f"Error in scheduled processing: {str(e)}")
        processing_stats["failed"] += 1
        raise

async def run_scheduler():
    """Run the scheduler in the background."""
    while True:
        schedule.run_pending()
        await asyncio.sleep(1)

@app.get("/health")
async def health_check():
    """Health check endpoint."""
    return {
        "status": "healthy",
        "timestamp": datetime.utcnow().isoformat(),
        "uptime": time.time()
    }

@app.get("/")
async def root():
    """Root endpoint with API information."""
    return {
        "name": "Email Extraction API",
        "version": "1.0.0",
        "status": "running",
        "timestamp": datetime.utcnow().isoformat(),
        "endpoints": {
            "health": "/health",
            "stats": "/stats",
            "accounts": "/accounts",
            "process": "/process-now",
            "process_account": "/process-account/{user_id}",
            "results": "/results",
            "iframe": "/iframe",
            "embed_code": "/embed-code",
            "docs": "/docs"
        }
    }

@app.get("/stats")
async def get_stats():
    """Get processing statistics."""
    return processing_stats

@app.get("/accounts")
async def get_accounts():
    """Get list of configured email accounts."""
    try:
        accounts = settings.get_email_accounts()
        return {
            "accounts": accounts,
            "total_count": len(accounts),
            "max_concurrent": settings.MAX_CONCURRENT_ACCOUNTS,
            "batch_size_per_account": settings.BATCH_SIZE_PER_ACCOUNT
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/process-account/{user_id}")
async def process_specific_account(user_id: str, batch_size: int = None):
    """Process emails for a specific account."""
    try:
        if user_id not in settings.get_email_accounts():
            raise HTTPException(status_code=404, detail=f"Account {user_id} not found in configuration")
        
        result = await multi_account_processor.process_account(user_id, batch_size)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/process-now")
async def process_now():
    """Trigger immediate processing of emails."""
    try:
        await process_emails()
        return {"status": "success", "message": "Processing completed"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/iframe", response_class=HTMLResponse)
async def iframe_interface():
    """Serve the iframe interface for embedding."""
    try:
        with open("static/index.html", "r") as f:
            html_content = f.read()
        return HTMLResponse(content=html_content)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="Interface not found")

@app.get("/embed-code", response_class=HTMLResponse)
async def get_embed_code():
    """Provide iframe embed code for integration."""
    embed_html = f'''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Email Processor - Embed Code</title>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 0;
                padding: 30px;
                background: #f5f7fa;
                line-height: 1.6;
            }}
            .container {{
                max-width: 800px;
                margin: 0 auto;
                background: white;
                padding: 40px;
                border-radius: 12px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            }}
            h1 {{
                color: #2d3748;
                margin-bottom: 20px;
                display: flex;
                align-items: center;
                gap: 10px;
            }}
            .code-block {{
                background: #1a202c;
                color: #e2e8f0;
                padding: 20px;
                border-radius: 8px;
                font-family: 'Monaco', 'Menlo', 'Ubuntu Mono', monospace;
                font-size: 14px;
                overflow-x: auto;
                margin: 20px 0;
                position: relative;
            }}
            .copy-btn {{
                position: absolute;
                top: 10px;
                right: 10px;
                background: #667eea;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                cursor: pointer;
                font-size: 12px;
            }}
            .copy-btn:hover {{
                background: #5a67d8;
            }}
            .info {{
                background: #e6fffa;
                border-left: 4px solid #38b2ac;
                padding: 15px;
                margin: 20px 0;
                border-radius: 4px;
            }}
            .btn {{
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                padding: 12px 24px;
                border: none;
                border-radius: 8px;
                text-decoration: none;
                display: inline-block;
                margin: 10px 10px 10px 0;
                font-weight: 600;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📋 Email Processor - Embed Code</h1>
            
            <h2>Basic Iframe Code</h2>
            <div class="code-block">
                <button class="copy-btn" onclick="copyToClipboard('basic-code')">Copy</button>
                <div id="basic-code">&lt;iframe 
    src="http://localhost:{settings.API_PORT}/iframe" 
    width="100%" 
    height="800" 
    frameborder="0" 
    style="border: 1px solid #ddd; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1);"
    allowfullscreen&gt;
&lt;/iframe&gt;</div>
            </div>

            <h2>Responsive Integration</h2>
            <div class="code-block">
                <button class="copy-btn" onclick="copyToClipboard('responsive-code')">Copy</button>
                <div id="responsive-code">&lt;div style="position: relative; width: 100%; height: 0; padding-bottom: 56.25%;"&gt;
    &lt;iframe 
        src="http://localhost:{settings.API_PORT}/iframe" 
        style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; border: none;"
        allowfullscreen&gt;
    &lt;/iframe&gt;
&lt;/div&gt;</div>
            </div>

            <div class="info">
                <strong>📝 Usage Instructions:</strong><br>
                1. Copy the iframe code above<br>
                2. Paste it into your HTML where you want the email processor to appear<br>
                3. For production, replace "localhost:{settings.API_PORT}" with your domain<br>
                4. Adjust width/height as needed for your layout
            </div>

            <a href="/iframe" class="btn" target="_blank">🔍 Preview Iframe</a>
            <a href="/docs" class="btn" target="_blank">📖 API Docs</a>
        </div>

        <script>
            function copyToClipboard(elementId) {{
                const element = document.getElementById(elementId);
                const text = element.textContent;
                navigator.clipboard.writeText(text).then(() => {{
                    alert('Code copied to clipboard!');
                }});
            }}
        </script>
    </body>
    </html>
    '''
    return HTMLResponse(content=embed_html)

@app.get("/favicon.ico")
async def favicon():
    """Serve favicon to prevent 404 errors."""
    # Return a simple 1x1 transparent PNG if no favicon exists
    return FileResponse("static/favicon.ico") if os.path.exists("static/favicon.ico") else HTMLResponse(content="", status_code=204)

@app.get("/results")
async def get_results():
    """Get recent extraction results."""
    try:
        # Get the most recent extraction file
        output_dir = "output"
        if not os.path.exists(output_dir):
            return {
                "message": "No results directory found",
                "results": [],
                "file": None
            }
        
        # Get all JSON files in output directory
        json_files = [f for f in os.listdir(output_dir) if f.endswith('.json')]
        
        if not json_files:
            return {
                "message": "No extraction results found",
                "results": [],
                "file": None
            }
        
        # Sort files by modification time, get the most recent
        json_files.sort(key=lambda f: os.path.getmtime(os.path.join(output_dir, f)), reverse=True)
        latest_file = json_files[0]
        
        # Read the latest results file
        with open(os.path.join(output_dir, latest_file), 'r') as f:
            results = json.load(f)
        
        return {
            "message": f"Showing {len(results)} recent extraction results",
            "results": results,
            "file": latest_file
        }
        
    except Exception as e:
        logger.error(f"Error retrieving results: {str(e)}")
        return {
            "message": "Error retrieving results",
            "results": [],
            "file": None,
            "error": str(e)
        }

if __name__ == "__main__":
    # Configure logging
    logger.add(
        "logs/email_processor.log",
        rotation="1 day",
        retention="7 days",
        level=settings.LOG_LEVEL
    )
    
    # Start the FastAPI server
    uvicorn.run(
        "main:app",
        host=settings.API_HOST,
        port=settings.API_PORT,
        reload=True
    ) 