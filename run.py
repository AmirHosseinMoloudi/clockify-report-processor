#!/usr/bin/env python3
"""
Clockify Report Processor - FastAPI Application Runner

This script starts the FastAPI application with proper configuration
for both development and production environments.
"""

import uvicorn
import os
import sys
from pathlib import Path

# Add the app directory to Python path
app_dir = Path(__file__).parent / "app"
sys.path.insert(0, str(app_dir))

def main():
    """Main entry point for the application."""
    
    # Configuration
    host = os.getenv("HOST", "127.0.0.1")
    port = int(os.getenv("PORT", "8000"))
    reload = os.getenv("RELOAD", "true").lower() == "true"
    log_level = os.getenv("LOG_LEVEL", "info")
    
    print(f"Starting Clockify Report Processor...")
    print(f"Server will be available at: http://{host}:{port}")
    print(f"Reload mode: {reload}")
    print(f"Log level: {log_level}")
    print("-" * 50)
    
    # Start the server
    uvicorn.run(
        "main:app",
        host=host,
        port=port,
        reload=reload,
        log_level=log_level,
        access_log=True
    )

if __name__ == "__main__":
    main()