import logging
import os
import traceback
from datetime import datetime
from typing import List, Optional

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

# Configure structured logging with enhanced format
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    handlers=[
        logging.FileHandler('clockify_processor.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Log application startup
logger.info("Initializing Clockify Report Processor v2.0.0")

# Initialize FastAPI app with nginx proxy configuration
app = FastAPI(
    title="Clockify Report Processor",
    description="Convert Clockify time reports into structured business reports",
    version="2.0.0",
    docs_url="/api/docs",
    redoc_url="/api/redoc",
    root_path="/clockify/report"  # Handle nginx proxy path
)

# Log FastAPI initialization
logger.info("FastAPI application initialized with root_path='/clockify/report'")

# Configure CORS for frontend access
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, specify exact origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static files at root for nginx proxy compatibility
app.mount("/static", StaticFiles(directory="static"), name="static")

# Ensure uploads directory exists
os.makedirs("uploads", exist_ok=True)
os.makedirs("exports", exist_ok=True)

# Global data storage (in production, use a database)
data_store = {}

class DataPreview(BaseModel):
    columns: List[str]
    rows: List[List[str]]
    total_records: int
    filename: str

class ExportRequest(BaseModel):
    export_type: str  # "projects" or "hr"
    filename: Optional[str] = None

@app.get("/", response_class=HTMLResponse)
async def read_root():
    """Serve the main application page"""
    try:
        with open("static/index.html", "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    except FileNotFoundError:
        logger.error("index.html not found in static directory")
        raise HTTPException(status_code=404, detail="Frontend not found")
    except Exception as e:
        logger.error(f"Error serving root page: {str(e)}")
        raise HTTPException(status_code=500, detail="Internal server error")

@app.get("/clockify/report/", response_class=HTMLResponse)
async def read_index_proxy():
    """Serve the main application page for nginx proxy path"""
    try:
        with open("static/index.html", "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    except FileNotFoundError:
        logger.error("index.html not found in static directory")
        raise HTTPException(status_code=404, detail="Frontend not found")
    except Exception as e:
        logger.error(f"Error serving root page: {str(e)}")
        raise HTTPException(status_code=500, detail="Internal server error")

@app.get("/clockify/report/api", response_class=HTMLResponse)
async def read_index_api():
    """Serve the main application page for API documentation access"""
    try:
        with open("static/index.html", "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    except FileNotFoundError:
        logger.error("index.html not found in static directory")
        raise HTTPException(status_code=404, detail="Frontend not found")
    except Exception as e:
        logger.error(f"Error serving root page: {str(e)}")
        raise HTTPException(status_code=500, detail="Internal server error")

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload and process Clockify Excel report with enhanced error handling"""
    start_time = datetime.now()
    logger.info(f"Upload request initiated - File: {file.filename}, Content-Type: {file.content_type}")
    
    try:
        # Enhanced file validation
        if not file.filename:
            logger.warning("Upload attempt with no filename")
            raise HTTPException(status_code=400, detail="No file provided")
            
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            logger.warning(f"Invalid file type uploaded: {file.filename}")
            raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are supported")
        
        # Check file size (limit to 50MB)
        content = await file.read()
        file_size_mb = len(content) / (1024 * 1024)
        if file_size_mb > 50:
            logger.warning(f"File too large: {file_size_mb:.2f}MB - {file.filename}")
            raise HTTPException(status_code=413, detail="File too large. Maximum size is 50MB")
        
        logger.info(f"File validation passed - Size: {file_size_mb:.2f}MB")
        
        # Save uploaded file with timestamp to avoid conflicts
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_filename = f"{timestamp}_{file.filename}"
        file_path = f"uploads/{safe_filename}"
        
        with open(file_path, "wb") as buffer:
            buffer.write(content)
        
        logger.info(f"File saved successfully: {file_path}")
        
        # Load and validate Excel file with enhanced error handling
        try:
            df = pd.read_excel(file_path)
            
            if df.empty:
                logger.error(f"Empty Excel file: {file.filename}")
                raise HTTPException(status_code=400, detail="Excel file is empty")
            
            # Validate required columns exist
            required_columns = ['Project', 'User', 'Description']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                logger.warning(f"Missing required columns: {missing_columns} in {file.filename}")
                # Don't fail, just log warning as column names may vary
            
            # Store data globally (in production, use session management)
            data_store['current_data'] = df
            data_store['filename'] = file.filename
            data_store['upload_time'] = start_time
            data_store['file_path'] = file_path
            
            processing_time = (datetime.now() - start_time).total_seconds()
            logger.info(f"Successfully processed {len(df)} records from {file.filename} in {processing_time:.2f}s")
            
            return {
                "message": "File uploaded successfully",
                "filename": file.filename,
                "records": len(df),
                "columns": list(df.columns),
                "processing_time_seconds": round(processing_time, 2),
                "file_size_mb": round(file_size_mb, 2)
            }
            
        except pd.errors.EmptyDataError:
            logger.error(f"Empty data in Excel file: {file.filename}")
            raise HTTPException(status_code=400, detail="Excel file contains no data")
        except pd.errors.ParserError as e:
            logger.error(f"Parser error reading Excel file {file.filename}: {str(e)}")
            raise HTTPException(status_code=400, detail=f"Invalid Excel file format: {str(e)}")
        except Exception as e:
            logger.error(f"Error reading Excel file {file.filename}: {str(e)}\n{traceback.format_exc()}")
            raise HTTPException(status_code=400, detail=f"Invalid Excel file: {str(e)}")
            
    except HTTPException:
        raise
    except Exception as e:
        processing_time = (datetime.now() - start_time).total_seconds()
        logger.error(f"Unexpected error in upload after {processing_time:.2f}s: {str(e)}\n{traceback.format_exc()}")
        raise HTTPException(status_code=500, detail="Internal server error during upload")

@app.get("/api/preview", response_model=DataPreview)
async def get_data_preview():
    """Get preview of uploaded data"""
    try:
        if 'current_data' not in data_store:
            raise HTTPException(status_code=404, detail="No data uploaded")
        
        df = data_store['current_data']
        filename = data_store.get('filename', 'unknown')
        
        # Get first 100 rows for preview
        preview_df = df.head(100)
        
        # Convert to serializable format
        rows = []
        for _, row in preview_df.iterrows():
            rows.append([str(val) if pd.notna(val) else "" for val in row])
        
        return DataPreview(
            columns=list(df.columns),
            rows=rows,
            total_records=len(df),
            filename=filename
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error generating preview: {str(e)}\n{traceback.format_exc()}")
        raise HTTPException(status_code=500, detail="Error generating data preview")

@app.post("/api/export")
async def export_data(request: ExportRequest):
    """Export data in specified format with enhanced error handling"""
    start_time = datetime.now()
    export_type = request.export_type.lower() if request.export_type else "unknown"
    logger.info(f"Export request initiated - Type: {export_type}, Custom filename: {request.filename}")
    
    try:
        # Validate data availability
        if 'current_data' not in data_store:
            logger.warning("Export attempted without uploaded data")
            raise HTTPException(status_code=404, detail="No data uploaded. Please upload a file first.")
        
        df = data_store['current_data'].copy()
        original_filename = data_store.get('filename', 'unknown')
        record_count = len(df)
        
        logger.info(f"Starting {export_type} export for {record_count} records from {original_filename}")
        
        # Validate export type
        valid_types = ["projects", "hr"]
        if export_type not in valid_types:
            logger.warning(f"Invalid export type requested: {export_type}")
            raise HTTPException(status_code=400, detail=f"Invalid export type '{export_type}'. Use 'projects' or 'hr'")
        
        # Generate export based on type
        try:
            if export_type == "projects":
                file_path = await export_projects_report(df, request.filename)
            elif export_type == "hr":
                file_path = await export_hr_report(df, request.filename)
            
            # Verify file was created successfully
            if not os.path.exists(file_path):
                logger.error(f"Export file was not created: {file_path}")
                raise HTTPException(status_code=500, detail="Export file generation failed")
            
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            processing_time = (datetime.now() - start_time).total_seconds()
            
            logger.info(f"Successfully exported {export_type} report: {file_path} ({file_size_mb:.2f}MB) in {processing_time:.2f}s")
            
            return {
                "message": f"{export_type.title()} report generated successfully",
                "download_url": f"/api/download/{os.path.basename(file_path)}",
                "filename": os.path.basename(file_path),
                "file_size_mb": round(file_size_mb, 2),
                "processing_time_seconds": round(processing_time, 2),
                "record_count": record_count
            }
            
        except Exception as export_error:
            logger.error(f"Error during {export_type} export generation: {str(export_error)}\n{traceback.format_exc()}")
            raise HTTPException(status_code=500, detail=f"Error generating {export_type} report: {str(export_error)}")
        
    except HTTPException:
        raise
    except Exception as e:
        processing_time = (datetime.now() - start_time).total_seconds()
        logger.error(f"Unexpected error during {export_type} export after {processing_time:.2f}s: {str(e)}\n{traceback.format_exc()}")
        raise HTTPException(status_code=500, detail="Internal server error during export")

@app.get("/api/download/{filename}")
async def download_file(filename: str):
    """Download exported file with enhanced security and logging"""
    logger.info(f"Download request for file: {filename}")
    
    try:
        # Sanitize filename to prevent directory traversal
        safe_filename = os.path.basename(filename)
        if safe_filename != filename:
            logger.warning(f"Potential directory traversal attempt: {filename}")
            raise HTTPException(status_code=400, detail="Invalid filename")
        
        file_path = f"exports/{safe_filename}"
        
        # Check if file exists
        if not os.path.exists(file_path):
            logger.warning(f"Requested file not found: {file_path}")
            raise HTTPException(status_code=404, detail="File not found")
        
        # Check file age (delete files older than 24 hours)
        file_age_hours = (datetime.now().timestamp() - os.path.getmtime(file_path)) / 3600
        if file_age_hours > 24:
            logger.info(f"Removing expired file: {file_path} (age: {file_age_hours:.1f}h)")
            os.remove(file_path)
            raise HTTPException(status_code=404, detail="File has expired and been removed")
        
        file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
        logger.info(f"Serving download: {safe_filename} ({file_size_mb:.2f}MB)")
        
        return FileResponse(
            path=file_path,
            filename=safe_filename,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={"Cache-Control": "no-cache, no-store, must-revalidate"}
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error downloading file {filename}: {str(e)}\n{traceback.format_exc()}")
        raise HTTPException(status_code=500, detail="Error downloading file")

async def export_projects_report(df: pd.DataFrame, custom_filename: Optional[str] = None) -> str:
    """Generate projects report with individual sheets for each project"""
    try:
        filename = custom_filename or f"projects_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = f"exports/{filename}"
        
        required_columns = [
            'Project', 'Description', 'User', 'Email', 
            'Start Date', 'Start Time', 'End Date', 'End Time', 'Duration (h)'
        ]
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # Get unique projects
            unique_projects = []
            for project in df['Project']:
                if project not in unique_projects and not pd.isna(project):
                    unique_projects.append(project)
            
            for project_name in unique_projects:
                # Filter data for this project
                project_data = df[df['Project'] == project_name].copy()
                
                # Create project dataframe
                project_df = pd.DataFrame(columns=required_columns)
                
                # Map existing columns
                for col in required_columns:
                    if col in project_data.columns:
                        project_df[col] = project_data[col]
                        
                        # Format dates
                        if col in ['Start Date', 'End Date'] and pd.api.types.is_datetime64_any_dtype(project_data[col]):
                            project_df[col] = project_data[col].dt.strftime('%d/%m/%Y')
                    else:
                        project_df[col] = None
                
                # Handle duration
                if 'Duration (h)' in project_data.columns:
                    project_df['Duration (h)'] = project_data['Duration (h)']
                elif 'Duration (decimal)' in project_data.columns:
                    project_df['Duration (h)'] = project_data['Duration (decimal)'].apply(decimal_to_time)
                
                # Calculate total duration
                total_seconds = calculate_total_duration(project_df['Duration (h)'])
                total_duration_str = seconds_to_time_string(total_seconds)
                
                # Add total row
                blank_row = pd.Series([None] * len(project_df.columns), index=project_df.columns)
                project_df = pd.concat([project_df, pd.DataFrame([blank_row])], ignore_index=True)
                
                total_row = pd.Series([None] * len(project_df.columns), index=project_df.columns)
                total_row['Project'] = 'Total:'
                total_row['Duration (h)'] = total_duration_str
                project_df = pd.concat([project_df, pd.DataFrame([total_row])], ignore_index=True)
                
                # Save to sheet
                sheet_name = sanitize_sheet_name(str(project_name))
                project_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        return file_path
        
    except Exception as e:
        logger.error(f"Error generating projects report: {str(e)}\n{traceback.format_exc()}")
        raise

async def export_hr_report(df: pd.DataFrame, custom_filename: Optional[str] = None) -> str:
    """Generate HR-friendly timesheet with individual sheets for each person"""
    try:
        filename = custom_filename or f"hr_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = f"exports/{filename}"
        
        required_columns = ['Project', 'Description', 'Time (h)']
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # Create sheets for each user
            if 'User' in df.columns:
                user_groups = df.groupby('User')
                
                for user_name, user_data in user_groups:
                    if pd.isna(user_name):
                        continue
                    
                    user_rows = []
                    user_total_seconds = 0
                    
                    # Group by project for this user
                    user_project_groups = user_data.groupby('Project')
                    
                    for project_name, project_data in user_project_groups:
                        if pd.isna(project_name):
                            continue
                        
                        # Calculate project total
                        project_seconds = calculate_total_duration_from_data(project_data)
                        project_time_str = seconds_to_time_string(project_seconds)
                        user_total_seconds += project_seconds
                        
                        # Add project row
                        user_rows.append({
                            'Project': project_name,
                            'Description': None,
                            'Time (h)': project_time_str
                        })
                        
                        # Group by description
                        desc_groups = {}
                        for _, row in project_data.iterrows():
                            desc = row.get('Description')
                            if pd.notna(desc):
                                entry_seconds = get_duration_seconds(row)
                                desc_groups[desc] = desc_groups.get(desc, 0) + entry_seconds
                        
                        # Add description rows
                        for desc, total_seconds in desc_groups.items():
                            total_time_str = seconds_to_time_string(total_seconds)
                            user_rows.append({
                                'Project': None,
                                'Description': desc,
                                'Time (h)': total_time_str
                            })
                    
                    # Create user dataframe
                    user_df = pd.DataFrame(user_rows)
                    
                    # Add total row
                    if not user_df.empty:
                        user_total_str = seconds_to_time_string(user_total_seconds)
                        
                        blank_row = pd.Series([None] * len(required_columns), index=required_columns)
                        user_df = pd.concat([user_df, pd.DataFrame([blank_row])], ignore_index=True)
                        
                        total_row = pd.Series([None] * len(required_columns), index=required_columns)
                        total_row['Project'] = 'Total:'
                        total_row['Time (h)'] = user_total_str
                        user_df = pd.concat([user_df, pd.DataFrame([total_row])], ignore_index=True)
                        
                        # Save user sheet
                        sheet_name = sanitize_sheet_name(str(user_name))
                        user_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        return file_path
        
    except Exception as e:
        logger.error(f"Error generating HR report: {str(e)}\n{traceback.format_exc()}")
        raise

def decimal_to_time(decimal_hours):
    """Convert decimal hours to HH:MM:SS format"""
    if pd.isna(decimal_hours):
        return None
    try:
        hours = int(decimal_hours)
        minutes = int((decimal_hours - hours) * 60)
        seconds = int(((decimal_hours - hours) * 60 - minutes) * 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    except (ValueError, TypeError):
        return "00:00:00"

def calculate_total_duration(duration_series):
    """Calculate total duration in seconds from a series of duration strings"""
    total_seconds = 0
    for duration in duration_series:
        if pd.notna(duration) and isinstance(duration, str):
            try:
                h, m, s = map(int, duration.split(':'))
                total_seconds += h * 3600 + m * 60 + s
            except (ValueError, AttributeError):
                pass
        elif pd.notna(duration) and hasattr(duration, 'hour'):
            total_seconds += duration.hour * 3600 + duration.minute * 60 + duration.second
    return total_seconds

def calculate_total_duration_from_data(project_data):
    """Calculate total duration from project data"""
    total_seconds = 0
    
    if 'Duration (decimal)' in project_data.columns:
        for hours in project_data['Duration (decimal)']:
            if pd.notna(hours):
                total_seconds += int(hours * 3600)
    elif 'Duration (h)' in project_data.columns:
        for duration in project_data['Duration (h)']:
            if pd.notna(duration) and isinstance(duration, str):
                try:
                    h, m, s = map(int, duration.split(':'))
                    total_seconds += h * 3600 + m * 60 + s
                except (ValueError, AttributeError):
                    pass
            elif pd.notna(duration) and hasattr(duration, 'hour'):
                total_seconds += duration.hour * 3600 + duration.minute * 60 + duration.second
    
    return total_seconds

def get_duration_seconds(row):
    """Get duration in seconds from a data row"""
    if 'Duration (decimal)' in row and pd.notna(row.get('Duration (decimal)')):
        return int(row.get('Duration (decimal)') * 3600)
    elif 'Duration (h)' in row and pd.notna(row.get('Duration (h)')):
        duration = row.get('Duration (h)')
        if isinstance(duration, str):
            try:
                h, m, s = map(int, duration.split(':'))
                return h * 3600 + m * 60 + s
            except (ValueError, AttributeError):
                pass
        elif hasattr(duration, 'hour'):
            return duration.hour * 3600 + duration.minute * 60 + duration.second
    return 0

def seconds_to_time_string(total_seconds):
    """Convert seconds to HH:MM:SS format"""
    hours = total_seconds // 3600
    remaining = total_seconds % 3600
    minutes = remaining // 60
    seconds = remaining % 60
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

def sanitize_sheet_name(name):
    """Sanitize sheet name for Excel compatibility"""
    return str(name)[:31].replace('/', '_').replace('\\', '_').replace('?', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_')

# Health check endpoint for monitoring
@app.get("/api/health")
async def health_check():
    """Health check endpoint for monitoring and load balancers"""
    try:
        # Check if directories exist
        uploads_exists = os.path.exists("uploads")
        exports_exists = os.path.exists("exports")
        
        # Check current data store status
        has_data = 'current_data' in data_store
        data_records = len(data_store.get('current_data', [])) if has_data else 0
        
        uptime_seconds = (datetime.now() - datetime.now()).total_seconds()  # This will be updated by startup event
        
        health_status = {
            "status": "healthy",
            "timestamp": datetime.now().isoformat(),
            "version": "2.0.0",
            "directories": {
                "uploads": uploads_exists,
                "exports": exports_exists
            },
            "data_store": {
                "has_data": has_data,
                "records": data_records
            },
            "uptime_seconds": uptime_seconds
        }
        
        logger.debug(f"Health check completed: {health_status}")
        return health_status
        
    except Exception as e:
        logger.error(f"Health check failed: {str(e)}")
        raise HTTPException(status_code=503, detail="Service unhealthy")

# Application startup event
@app.on_event("startup")
async def startup_event():
    """Initialize application on startup"""
    logger.info("Application startup initiated")
    
    # Ensure required directories exist
    os.makedirs("uploads", exist_ok=True)
    os.makedirs("exports", exist_ok=True)
    
    # Clean up old files on startup
    cleanup_old_files()
    
    logger.info("Application startup completed successfully")

def cleanup_old_files():
    """Clean up files older than 24 hours"""
    try:
        current_time = datetime.now().timestamp()
        cleaned_count = 0
        
        for directory in ["uploads", "exports"]:
            if os.path.exists(directory):
                for filename in os.listdir(directory):
                    file_path = os.path.join(directory, filename)
                    if os.path.isfile(file_path):
                        file_age_hours = (current_time - os.path.getmtime(file_path)) / 3600
                        if file_age_hours > 24:
                            os.remove(file_path)
                            cleaned_count += 1
                            logger.info(f"Cleaned up old file: {file_path}")
        
        if cleaned_count > 0:
            logger.info(f"Startup cleanup completed: {cleaned_count} files removed")
        else:
            logger.info("Startup cleanup completed: no old files found")
            
    except Exception as e:
        logger.error(f"Error during startup cleanup: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    logger.info("Starting application in direct mode")
    uvicorn.run(app, host="0.0.0.0", port=8002, log_level="info")