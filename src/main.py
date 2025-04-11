import sys
import os
import signal
import pandas as pd
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QLineEdit,
                             QListWidget, QListWidgetItem, QFrame, QSplitter,
                             QMessageBox, QSizePolicy, QFileDialog, QProgressBar,
                             QStatusBar, QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QFont

class ResponsiveApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Clockify Report Processor")
        self.resize(1000, 600)
        self.setMinimumSize(600, 400)
        
        # Set up the central widget
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # Main layout
        self.main_layout = QHBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        
        # Create sidebar
        self.create_sidebar()
        
        # Create content area
        self.create_content()
        
        # Add splitter to make the UI responsive
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.addWidget(self.sidebar)
        self.splitter.addWidget(self.content_area)
        self.splitter.setSizes([200, 800])
        
        self.main_layout.addWidget(self.splitter)
        
        # Create status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")
        
        # Initialize data variables
        self.clockify_data = None
        self.input_file_path = None
        
    def create_sidebar(self):
        # Sidebar
        self.sidebar = QWidget()
        self.sidebar.setObjectName("sidebar")
        self.sidebar.setStyleSheet("""
            #sidebar {
                background-color: #2E3440;
                min-width: 180px;
                max-width: 250px;
            }
            QPushButton {
                text-align: left;
                padding: 10px;
                border: none;
                border-radius: 0px;
                color: #ECEFF4;
                font-size: 14px;
                background-color: transparent;
            }
            QPushButton:hover {
                background-color: #3B4252;
            }
            QPushButton:pressed {
                background-color: #4C566A;
            }
            QLabel {
                color: #ECEFF4;
                font-size: 18px;
                font-weight: bold;
                padding: 15px 10px;
            }
        """)
        
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(0)
        
        # App title
        title = QLabel("Clockify Processor")
        title.setAlignment(Qt.AlignCenter)
        sidebar_layout.addWidget(title)
        
        # Separator
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: #3B4252; max-height: 1px;")
        sidebar_layout.addWidget(separator)
        
        # Navigation buttons
        nav_buttons = [
            {"name": "Import", "action": self.import_excel},
            {"name": "View Data", "action": self.view_data},
            {"name": "Export Projects", "action": self.export_projects},
            {"name": "Export HR", "action": self.export_hr},
            {"name": "Settings", "action": None}
        ]
        
        for btn in nav_buttons:
            button = QPushButton(f"  {btn['name']}")
            if btn['action']:
                button.clicked.connect(btn['action'])
            sidebar_layout.addWidget(button)
        
        sidebar_layout.addStretch()
        
        # Logout button at the bottom
        logout_btn = QPushButton("  Exit")
        logout_btn.clicked.connect(self.close)
        sidebar_layout.addWidget(logout_btn)
        
    def create_content(self):
        # Content area
        self.content_area = QWidget()
        self.content_area.setObjectName("content")
        self.content_area.setStyleSheet("""
            #content {
                background-color: #ECEFF4;
            }
            QLabel {
                color: #2E3440;
            }
            QPushButton {
                background-color: #5E81AC;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 15px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #81A1C1;
            }
            QPushButton:pressed {
                background-color: #4C566A;
            }
            QLineEdit {
                border: 1px solid #D8DEE9;
                border-radius: 4px;
                padding: 8px;
                background-color: white;
            }
            QTableWidget {
                border: 1px solid #D8DEE9;
                border-radius: 4px;
                background-color: white;
            }
            QProgressBar {
                border: 1px solid #D8DEE9;
                border-radius: 5px;
                background-color: #ECEFF4;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #5E81AC;
                border-radius: 5px;
            }
        """)
        
        self.content_layout = QVBoxLayout(self.content_area)
        self.content_layout.setContentsMargins(20, 20, 20, 20)
        self.content_layout.setSpacing(15)
        
        # Welcome message and instructions
        self.welcome_widget = QWidget()
        welcome_layout = QVBoxLayout(self.welcome_widget)
        
        header_label = QLabel("Clockify Report Processor")
        header_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        welcome_layout.addWidget(header_label)
        
        instructions = QLabel(
            "This application processes Clockify time reports and exports them into two formats:\n"
            "1. projects.xlsx - Project-based summary with detailed time tracking\n"
            "2. hr.xlsx - HR-friendly timesheet with project descriptions and hours\n\n"
            "To get started, click 'Import' and select your Clockify Excel report."
        )
        instructions.setStyleSheet("font-size: 14px; line-height: 1.4;")
        instructions.setWordWrap(True)
        welcome_layout.addWidget(instructions)
        
        import_button = QPushButton("Import Clockify Report")
        import_button.setFixedWidth(200)
        import_button.clicked.connect(self.import_excel)
        welcome_layout.addWidget(import_button)
        
        welcome_layout.addStretch()
        
        self.content_layout.addWidget(self.welcome_widget)
        
        # Table widget for data preview (initially hidden)
        self.table_widget = QTableWidget()
        self.table_widget.setHidden(True)
        self.content_layout.addWidget(self.table_widget)
        
        # Progress bar (initially hidden)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setHidden(True)
        self.content_layout.addWidget(self.progress_bar)
        
        # Export buttons layout (initially hidden)
        self.export_widget = QWidget()
        self.export_layout = QHBoxLayout(self.export_widget)
        
        self.export_projects_btn = QPushButton("Export Projects")
        self.export_projects_btn.clicked.connect(self.export_projects)
        self.export_layout.addWidget(self.export_projects_btn)
        
        self.export_hr_btn = QPushButton("Export HR")
        self.export_hr_btn.clicked.connect(self.export_hr)
        self.export_layout.addWidget(self.export_hr_btn)
        
        self.export_widget.setHidden(True)
        self.content_layout.addWidget(self.export_widget)
    
    def import_excel(self):
        """Import a Clockify Excel report file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Select Clockify Report", 
            "", 
            "Excel Files (*.xlsx *.xls)"
        )
        
        if not file_path:
            return
            
        self.input_file_path = file_path
        self.status_bar.showMessage(f"Loading file: {os.path.basename(file_path)}")
        
        try:
            # Show progress
            self.progress_bar.setHidden(False)
            self.progress_bar.setValue(25)
            
            # Load the Excel file
            self.clockify_data = pd.read_excel(file_path)
            
            # Update progress
            self.progress_bar.setValue(100)
            
            # Show data preview
            self.display_data_preview()
            
            # Show export buttons
            self.export_widget.setHidden(False)
            
            # Update status
            self.status_bar.showMessage(f"Loaded {len(self.clockify_data)} records from {os.path.basename(file_path)}")
            
        except Exception as e:
            self.progress_bar.setHidden(True)
            QMessageBox.critical(self, "Error", f"Failed to load file: {str(e)}")
            self.status_bar.showMessage("Error loading file")
    
    def display_data_preview(self):
        """Display a preview of the loaded data in the table widget"""
        if self.clockify_data is None:
            return
            
        # Hide welcome widget and show table
        self.welcome_widget.setHidden(True)
        self.table_widget.setHidden(False)
        
        # Get the first 100 rows for preview
        preview_data = self.clockify_data.head(100)
        
        # Set up table
        self.table_widget.setRowCount(len(preview_data))
        self.table_widget.setColumnCount(len(preview_data.columns))
        self.table_widget.setHorizontalHeaderLabels(preview_data.columns)
        
        # Populate table
        for row in range(len(preview_data)):
            for col in range(len(preview_data.columns)):
                value = str(preview_data.iloc[row, col])
                item = QTableWidgetItem(value)
                self.table_widget.setItem(row, col, item)
        
        # Resize columns to contents
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        
        # Add data summary
        summary_label = QLabel(f"Loaded {len(self.clockify_data)} time entries from {os.path.basename(self.input_file_path)}")
        summary_label.setStyleSheet("font-size: 14px;")
        self.status_bar.showMessage(summary_label.text())
    
    def view_data(self):
        """Switch to data view"""
        if self.clockify_data is not None:
            self.welcome_widget.setHidden(True)
            self.table_widget.setHidden(False)
            self.export_widget.setHidden(False)
    
    def export_projects(self):
        """Export project-based summary to projects.xlsx with dedicated sheets for each project"""
        if self.clockify_data is None:
            QMessageBox.warning(self, "No Data", "Please import a Clockify report first.")
            return
            
        try:
            # Get save file location
            file_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Save Projects Report", 
                "projects.xlsx", 
                "Excel Files (*.xlsx)"
            )
            
            if not file_path:
                return
                
            self.status_bar.showMessage("Processing projects data...")
            self.progress_bar.setHidden(False)
            self.progress_bar.setValue(10)
            
            # Make a copy of the dataframe to work with
            df = self.clockify_data.copy()
            
            # Required columns for the projects.xlsx format
            required_columns = [
                'Project', 'Description', 'User', 'Email', 
                'Start Date', 'Start Time', 'End Date', 'End Time', 'Duration (h)'
            ]
            
            # Create a writer to save multiple sheets
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Create a main sheet with all projects
                main_df = pd.DataFrame(columns=required_columns)
                
                # Map existing columns to the target format
                if 'Project' in df.columns:
                    main_df['Project'] = df['Project']
                else:
                    main_df['Project'] = None
                
                if 'Description' in df.columns:
                    main_df['Description'] = df['Description']
                else:
                    main_df['Description'] = None
                
                if 'User' in df.columns:
                    main_df['User'] = df['User']
                else:
                    main_df['User'] = None
                
                if 'Email' in df.columns:
                    main_df['Email'] = df['Email']
                else:
                    main_df['Email'] = None
                
                # Handle date and time columns
                if 'Start Date' in df.columns:
                    main_df['Start Date'] = df['Start Date']
                    if pd.api.types.is_datetime64_any_dtype(df['Start Date']):
                        main_df['Start Date'] = df['Start Date'].dt.strftime('%d/%m/%Y')
                else:
                    main_df['Start Date'] = None
                
                if 'Start Time' in df.columns:
                    main_df['Start Time'] = df['Start Time']
                else:
                    main_df['Start Time'] = None
                
                if 'End Date' in df.columns:
                    main_df['End Date'] = df['End Date']
                    if pd.api.types.is_datetime64_any_dtype(df['End Date']):
                        main_df['End Date'] = df['End Date'].dt.strftime('%d/%m/%Y')
                else:
                    main_df['End Date'] = None
                
                if 'End Time' in df.columns:
                    main_df['End Time'] = df['End Time']
                else:
                    main_df['End Time'] = None
                
                # Handle duration
                if 'Duration (h)' in df.columns:
                    main_df['Duration (h)'] = df['Duration (h)']
                elif 'Duration (decimal)' in df.columns:
                    # Convert decimal hours to HH:MM:SS format
                    def decimal_to_time(decimal_hours):
                        if pd.isna(decimal_hours):
                            return None
                        hours = int(decimal_hours)
                        minutes = int((decimal_hours - hours) * 60)
                        seconds = int(((decimal_hours - hours) * 60 - minutes) * 60)
                        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                    
                    main_df['Duration (h)'] = df['Duration (decimal)'].apply(decimal_to_time)
                else:
                    # Try to extract duration from time entries if available
                    main_df['Duration (h)'] = "00:00:00"
                
                self.progress_bar.setValue(30)
                
                # Calculate total duration for all entries
                total_duration_seconds = 0
                
                # Process durations in HH:MM:SS format
                for duration in main_df['Duration (h)']:
                    if pd.notna(duration) and isinstance(duration, str):
                        try:
                            h, m, s = map(int, duration.split(':'))
                            total_duration_seconds += h * 3600 + m * 60 + s
                        except (ValueError, AttributeError):
                            pass
                    elif pd.notna(duration) and hasattr(duration, 'hour'):
                        # Handle time objects
                        total_duration_seconds += duration.hour * 3600 + duration.minute * 60 + duration.second
                
                # Convert total seconds to HH:MM:SS
                total_hours = total_duration_seconds // 3600
                remaining_seconds = total_duration_seconds % 3600
                total_minutes = remaining_seconds // 60
                total_seconds = remaining_seconds % 60
                total_duration_str = f"{total_hours:02d}:{total_minutes:02d}:{total_seconds:02d}"
                
                # Add total row to main sheet
                blank_row = pd.Series([None] * len(main_df.columns), index=main_df.columns)
                main_df = pd.concat([main_df, pd.DataFrame([blank_row])], ignore_index=True)
                
                total_row = pd.Series([None] * len(main_df.columns), index=main_df.columns)
                total_row['Project'] = 'Total:'
                total_row['Duration (h)'] = total_duration_str
                main_df = pd.concat([main_df, pd.DataFrame([total_row])], ignore_index=True)
                
                # Skip creating the 'All Projects' sheet as per user request
                self.progress_bar.setValue(50)
                
                # Group by Project and create a sheet for each project
                # Get unique project names while preserving order
                unique_projects = []
                for project in df['Project']:
                    if project not in unique_projects and not pd.isna(project):
                        unique_projects.append(project)
                
                project_count = len(unique_projects)
                current_project = 0
                
                for project_name in unique_projects:
                    current_project += 1
                    progress = 50 + (current_project / project_count) * 40
                    self.progress_bar.setValue(int(progress))
                    
                    # Filter data for this project without groupby to preserve duplicates
                    project_data = df[df['Project'] == project_name].copy()
                    
                    # Create a dataframe for this project
                    project_df = pd.DataFrame(columns=required_columns)
                    
                    # Fill in the data for this project (preserve all entries including duplicates)
                    for col in required_columns:
                        if col in project_data.columns:
                            project_df[col] = project_data[col]
                            
                            # Format dates if needed
                            if col in ['Start Date', 'End Date'] and pd.api.types.is_datetime64_any_dtype(project_data[col]):
                                project_df[col] = project_data[col].dt.strftime('%d/%m/%Y')
                        else:
                            project_df[col] = None
                    
                    # Handle duration for this project
                    if 'Duration (h)' in project_data.columns:
                        project_df['Duration (h)'] = project_data['Duration (h)']
                    elif 'Duration (decimal)' in project_data.columns:
                        project_df['Duration (h)'] = project_data['Duration (decimal)'].apply(decimal_to_time)
                    
                    # Calculate total duration for this project
                    project_duration_seconds = 0
                    for duration in project_df['Duration (h)']:
                        if pd.notna(duration) and isinstance(duration, str):
                            try:
                                h, m, s = map(int, duration.split(':'))
                                project_duration_seconds += h * 3600 + m * 60 + s
                            except (ValueError, AttributeError):
                                pass
                        elif pd.notna(duration) and hasattr(duration, 'hour'):
                            project_duration_seconds += duration.hour * 3600 + duration.minute * 60 + duration.second
                    
                    # Convert project total seconds to HH:MM:SS
                    project_hours = project_duration_seconds // 3600
                    remaining = project_duration_seconds % 3600
                    project_minutes = remaining // 60
                    project_seconds = remaining % 60
                    project_duration_str = f"{project_hours:02d}:{project_minutes:02d}:{project_seconds:02d}"
                    
                    # Add total row to project sheet
                    blank_row = pd.Series([None] * len(project_df.columns), index=project_df.columns)
                    project_df = pd.concat([project_df, pd.DataFrame([blank_row])], ignore_index=True)
                    
                    total_row = pd.Series([None] * len(project_df.columns), index=project_df.columns)
                    total_row['Project'] = 'Total:'
                    total_row['Duration (h)'] = project_duration_str
                    project_df = pd.concat([project_df, pd.DataFrame([total_row])], ignore_index=True)
                    
                    # Save project sheet - use a valid sheet name (max 31 chars, no special chars)
                    sheet_name = str(project_name)[:31].replace('/', '_').replace('\\', '_').replace('?', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_')
                    project_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                self.progress_bar.setValue(95)
            
            self.progress_bar.setValue(100)
            self.status_bar.showMessage(f"Projects report saved to {file_path}")
            
            QMessageBox.information(self, "Export Complete", f"Projects report exported to {file_path} with individual sheets for each project")
            
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export projects report: {str(e)}")
            self.status_bar.showMessage("Export failed")
        finally:
            self.progress_bar.setHidden(True)
    
    def export_hr(self):
        """Export HR-friendly timesheet to hr.xlsx with dedicated sheets for each person"""
        if self.clockify_data is None:
            QMessageBox.warning(self, "No Data", "Please import a Clockify report first.")
            return
            
        try:
            # Get save file location
            file_path, _ = QFileDialog.getSaveFileName(
                self, 
                "Save HR Report", 
                "hr.xlsx", 
                "Excel Files (*.xlsx)"
            )
            
            if not file_path:
                return
                
            self.status_bar.showMessage("Processing HR data...")
            self.progress_bar.setHidden(False)
            self.progress_bar.setValue(10)
            
            # Make a copy of the dataframe to work with
            df = self.clockify_data.copy()
            
            # Required columns for HR format
            required_columns = ['Project', 'Description', 'Time (h)']
            
            # Create a writer to save multiple sheets
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Create a main sheet with all entries
                all_rows = []
                grand_total_seconds = 0
                
                # Group by Project to get the main entries
                project_groups = df.groupby('Project')
                
                for project_name, project_data in project_groups:
                    if pd.isna(project_name):
                        continue
                        
                    # Calculate total time for this project
                    project_seconds = 0
                    
                    # Try to get duration from different possible columns
                    if 'Duration (decimal)' in project_data.columns:
                        # Convert decimal hours to seconds
                        for hours in project_data['Duration (decimal)']:
                            if pd.notna(hours):
                                project_seconds += int(hours * 3600)
                    elif 'Duration (h)' in project_data.columns:
                        # Parse HH:MM:SS format
                        for duration in project_data['Duration (h)']:
                            if pd.notna(duration) and isinstance(duration, str):
                                try:
                                    h, m, s = map(int, duration.split(':'))
                                    project_seconds += h * 3600 + m * 60 + s
                                except (ValueError, AttributeError):
                                    pass
                            elif pd.notna(duration) and hasattr(duration, 'hour'):
                                # Handle time objects
                                project_seconds += duration.hour * 3600 + duration.minute * 60 + duration.second
                    
                    # Convert project seconds to HH:MM:SS
                    project_hours = project_seconds // 3600
                    remaining = project_seconds % 3600
                    project_minutes = remaining // 60
                    project_seconds_remainder = remaining % 60
                    project_time_str = f"{project_hours:02d}:{project_minutes:02d}:{project_seconds_remainder:02d}"
                    
                    # Add to grand total
                    grand_total_seconds += project_seconds
                    
                    # Add the main project row
                    project_row = {
                        'Project': project_name,
                        'Description': None,
                        'Time (h)': project_time_str
                    }
                    all_rows.append(project_row)
                    
                    # Group entries by description and sum their durations
                    desc_groups = {}
                    
                    for idx, row in project_data.iterrows():
                        desc = row.get('Description')
                        if pd.notna(desc):  # Only process if description is not NA
                            # Get duration for this individual entry
                            entry_seconds = 0
                            
                            if 'Duration (decimal)' in project_data.columns and pd.notna(row.get('Duration (decimal)')):
                                entry_seconds = int(row.get('Duration (decimal)') * 3600)
                            elif 'Duration (h)' in project_data.columns and pd.notna(row.get('Duration (h)')):
                                duration = row.get('Duration (h)')
                                if isinstance(duration, str):
                                    try:
                                        h, m, s = map(int, duration.split(':'))
                                        entry_seconds = h * 3600 + m * 60 + s
                                    except (ValueError, AttributeError):
                                        pass
                                elif hasattr(duration, 'hour'):
                                    entry_seconds = duration.hour * 3600 + duration.minute * 60 + duration.second
                            
                            # Add to the description group total
                            if desc in desc_groups:
                                desc_groups[desc] += entry_seconds
                            else:
                                desc_groups[desc] = entry_seconds
                    
                    # Create a row for each unique description with summed duration
                    for desc, total_seconds in desc_groups.items():
                        # Convert total seconds to HH:MM:SS
                        total_hours = total_seconds // 3600
                        remaining = total_seconds % 3600
                        total_minutes = remaining // 60
                        total_seconds_remainder = remaining % 60
                        total_time_str = f"{total_hours:02d}:{total_minutes:02d}:{total_seconds_remainder:02d}"
                        
                        desc_row = {
                            'Project': None,
                            'Description': desc,
                            'Time (h)': total_time_str
                        }
                        all_rows.append(desc_row)
                
                # Create the DataFrame from all rows
                hr_df = pd.DataFrame(all_rows)
                
                self.progress_bar.setValue(40)
                
                # Get date range for the total row
                start_date = None
                end_date = None
                if 'Start Date' in df.columns and not df['Start Date'].empty:
                    if pd.api.types.is_datetime64_any_dtype(df['Start Date']):
                        start_date = df['Start Date'].min().strftime('%d/%m/%Y')
                    else:
                        # Try to parse the date strings
                        try:
                            dates = pd.to_datetime(df['Start Date'])
                            start_date = dates.min().strftime('%d/%m/%Y')
                        except:
                            pass
                
                if 'End Date' in df.columns and not df['End Date'].empty:
                    if pd.api.types.is_datetime64_any_dtype(df['End Date']):
                        end_date = df['End Date'].max().strftime('%d/%m/%Y')
                    else:
                        # Try to parse the date strings
                        try:
                            dates = pd.to_datetime(df['End Date'])
                            end_date = dates.max().strftime('%d/%m/%Y')
                        except:
                            pass
                
                # Convert grand total seconds to HH:MM:SS
                grand_total_hours = grand_total_seconds // 3600
                remaining = grand_total_seconds % 3600
                grand_total_minutes = remaining // 60
                grand_total_seconds_remainder = remaining % 60
                grand_total_str = f"{grand_total_hours:02d}:{grand_total_minutes:02d}:{grand_total_seconds_remainder:02d}"
                
                # Add total row with date range if available
                date_range = ""
                if start_date and end_date:
                    date_range = f"Total ({start_date} - {end_date})"
                else:
                    date_range = "Total"
                
                # Add blank row before total
                blank_row = pd.Series([None] * len(required_columns), index=required_columns)
                hr_df = pd.concat([hr_df, pd.DataFrame([blank_row])], ignore_index=True)
                
                # Add total row
                total_row = pd.Series([None] * len(required_columns), index=required_columns)
                total_row['Project'] = date_range
                total_row['Time (h)'] = f"Total:\n{grand_total_str}"
                hr_df = pd.concat([hr_df, pd.DataFrame([total_row])], ignore_index=True)
                
                # Skip creating the 'All Entries' sheet as per user request
                self.progress_bar.setValue(50)
                
                # Create sheets for each person
                if 'User' in df.columns:
                    user_groups = df.groupby('User')
                    user_count = len(user_groups)
                    current_user = 0
                    
                    for user_name, user_data in user_groups:
                        if pd.isna(user_name):
                            continue
                        
                        current_user += 1
                        progress = 50 + (current_user / user_count) * 40
                        self.progress_bar.setValue(int(progress))
                        
                        # Create a dataframe for this user
                        user_rows = []
                        user_total_seconds = 0
                        
                        # Group by Project for this user
                        user_project_groups = user_data.groupby('Project')
                        
                        for project_name, project_data in user_project_groups:
                            if pd.isna(project_name):
                                continue
                            
                            # Calculate total time for this project
                            project_seconds = 0
                            
                            # Try to get duration from different possible columns
                            if 'Duration (decimal)' in project_data.columns:
                                for hours in project_data['Duration (decimal)']:
                                    if pd.notna(hours):
                                        project_seconds += int(hours * 3600)
                            elif 'Duration (h)' in project_data.columns:
                                for duration in project_data['Duration (h)']:
                                    if pd.notna(duration) and isinstance(duration, str):
                                        try:
                                            h, m, s = map(int, duration.split(':'))
                                            project_seconds += h * 3600 + m * 60 + s
                                        except (ValueError, AttributeError):
                                            pass
                                    elif pd.notna(duration) and hasattr(duration, 'hour'):
                                        project_seconds += duration.hour * 3600 + duration.minute * 60 + duration.second
                            
                            # Convert project seconds to HH:MM:SS
                            project_hours = project_seconds // 3600
                            remaining = project_seconds % 3600
                            project_minutes = remaining // 60
                            project_seconds_remainder = remaining % 60
                            project_time_str = f"{project_hours:02d}:{project_minutes:02d}:{project_seconds_remainder:02d}"
                            
                            # Add to user total
                            user_total_seconds += project_seconds
                            
                            # Add the main project row
                            project_row = {
                                'Project': project_name,
                                'Description': None,
                                'Time (h)': project_time_str
                            }
                            user_rows.append(project_row)
                            
                            # Group entries by description and sum their durations
                            desc_groups = {}
                            
                            for idx, row in project_data.iterrows():
                                desc = row.get('Description')
                                if pd.notna(desc):  # Only process if description is not NA
                                    # Get duration for this individual entry
                                    entry_seconds = 0
                                    
                                    if 'Duration (decimal)' in project_data.columns and pd.notna(row.get('Duration (decimal)')):
                                        entry_seconds = int(row.get('Duration (decimal)') * 3600)
                                    elif 'Duration (h)' in project_data.columns and pd.notna(row.get('Duration (h)')):
                                        duration = row.get('Duration (h)')
                                        if isinstance(duration, str):
                                            try:
                                                h, m, s = map(int, duration.split(':'))
                                                entry_seconds = h * 3600 + m * 60 + s
                                            except (ValueError, AttributeError):
                                                pass
                                        elif hasattr(duration, 'hour'):
                                            entry_seconds = duration.hour * 3600 + duration.minute * 60 + duration.second
                                    
                                    # Add to the description group total
                                    if desc in desc_groups:
                                        desc_groups[desc] += entry_seconds
                                    else:
                                        desc_groups[desc] = entry_seconds
                            
                            # Create a row for each unique description with summed duration
                            for desc, total_seconds in desc_groups.items():
                                # Convert total seconds to HH:MM:SS
                                total_hours = total_seconds // 3600
                                remaining = total_seconds % 3600
                                total_minutes = remaining // 60
                                total_seconds_remainder = remaining % 60
                                total_time_str = f"{total_hours:02d}:{total_minutes:02d}:{total_seconds_remainder:02d}"
                                
                                desc_row = {
                                    'Project': None,
                                    'Description': desc,
                                    'Time (h)': total_time_str
                                }
                                user_rows.append(desc_row)
                        
                        # Create the DataFrame for this user
                        user_df = pd.DataFrame(user_rows)
                        
                        # Convert user total seconds to HH:MM:SS
                        user_total_hours = user_total_seconds // 3600
                        remaining = user_total_seconds % 3600
                        user_total_minutes = remaining // 60
                        user_total_seconds_remainder = remaining % 60
                        user_total_str = f"{user_total_hours:02d}:{user_total_minutes:02d}:{user_total_seconds_remainder:02d}"
                        
                        # Add blank row before total
                        if not user_df.empty:
                            blank_row = pd.Series([None] * len(required_columns), index=required_columns)
                            user_df = pd.concat([user_df, pd.DataFrame([blank_row])], ignore_index=True)
                            
                            # Add total row with date range if available
                            date_range = ""
                            if start_date and end_date:
                                date_range = f"Total ({start_date} - {end_date})"
                            else:
                                date_range = "Total"
                            
                            total_row = pd.Series([None] * len(required_columns), index=required_columns)
                            total_row['Project'] = date_range
                            total_row['Time (h)'] = f"Total:\n{user_total_str}"
                            user_df = pd.concat([user_df, pd.DataFrame([total_row])], ignore_index=True)
                            
                            # Save user sheet - use a valid sheet name (max 31 chars, no special chars)
                            sheet_name = str(user_name)[:31].replace('/', '_').replace('\\', '_').replace('?', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_')
                            user_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                self.progress_bar.setValue(95)
            
            self.progress_bar.setValue(100)
            self.status_bar.showMessage(f"HR report saved to {file_path}")
            
            QMessageBox.information(self, "Export Complete", f"HR report exported to {file_path} with individual sheets for each person")
            
            self.progress_bar.setValue(100)
            self.status_bar.showMessage(f"HR report saved to {file_path}")
            
            QMessageBox.information(self, "Export Complete", f"HR report exported to {file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export HR report: {str(e)}")
            self.status_bar.showMessage("Export failed")
        finally:
            self.progress_bar.setHidden(True)

def signal_handler(sig, frame):
    """Handle Ctrl+C signal"""
    print("\nExiting application gracefully...")
    QApplication.quit()
    sys.exit(0)

def main():
    # Set up signal handler for clean termination
    signal.signal(signal.SIGINT, signal_handler)
    
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    # Set global font
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    window = ResponsiveApp()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()