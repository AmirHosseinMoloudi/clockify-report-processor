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

# Configure structured logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('clockify_processor.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Initialize FastAPI app
app = FastAPI(
    title="Clockify Report Processor",
    description="Convert Clockify time reports into structured business reports",
    version="2.0.0",
    docs_url="/api/docs",
    redoc_url="/api/redoc"
)

# Configure CORS for frontend access
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, specify exact origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static files
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

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload and process Clockify Excel report"""
    try:
        logger.info(f"Processing upload: {file.filename}")
        
        # Validate file type
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are supported")
        
        # Save uploaded file
        file_path = f"uploads/{file.filename}"
        with open(file_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)
        
        # Load and validate Excel file
        try:
            df = pd.read_excel(file_path)
            if df.empty:
                raise HTTPException(status_code=400, detail="Excel file is empty")
            
            # Store data globally (in production, use session management)
            data_store['current_data'] = df
            data_store['filename'] = file.filename
            
            logger.info(f"Successfully loaded {len(df)} records from {file.filename}")
            
            return {
                "message": "File uploaded successfully",
                "filename": file.filename,
                "records": len(df),
                "columns": list(df.columns)
            }
            
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            raise HTTPException(status_code=400, detail=f"Invalid Excel file: {str(e)}")
            
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Unexpected error in upload: {str(e)}\n{traceback.format_exc()}")
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
    """Export data in specified format"""
    try:
        if 'current_data' not in data_store:
            raise HTTPException(status_code=404, detail="No data uploaded")
        
        df = data_store['current_data'].copy()
        export_type = request.export_type.lower()
        
        if export_type == "projects":
            file_path = await export_projects_report(df, request.filename)
        elif export_type == "hr":
            file_path = await export_hr_report(df, request.filename)
        else:
            raise HTTPException(status_code=400, detail="Invalid export type. Use 'projects' or 'hr'")
        
        logger.info(f"Successfully exported {export_type} report to {file_path}")
        
        return {
            "message": f"{export_type.title()} report generated successfully",
            "download_url": f"/api/download/{os.path.basename(file_path)}"
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error during export: {str(e)}\n{traceback.format_exc()}")
        raise HTTPException(status_code=500, detail="Error generating export")

@app.get("/api/download/{filename}")
async def download_file(filename: str):
    """Download exported file"""
    try:
        file_path = f"exports/{filename}"
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="File not found")
        
        return FileResponse(
            path=file_path,
            filename=filename,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error downloading file: {str(e)}")
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

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")