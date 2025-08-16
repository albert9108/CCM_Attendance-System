# CCM Attendance System

A comprehensive QR code-based attendance tracking system designed for educational institutions. This system allows students to check in and out using QR codes, maintaining accurate attendance records on a daily and monthly basis.

![CCM Attendance System](CCMLogo.png)

## Features

### Core Features
- **QR Code Scanning**: Fast and efficient check-in/check-out with QR code scanning
- **Real-time Attendance Tracking**: View attendance status instantly with live updates
- **Daily and Monthly Reports**: Automatically generates Excel reports for day and month attendance
- **Student Database Integration**: Works with your existing student database
- **User-friendly Interface**: Clean, intuitive interface with clear visual indicators
- **Attendance Statistics**: Tracks attendance rates for each student
- **Backup System**: Automatic backup of attendance records

### New Features (v1.7)
- **📁 Database File Selection**: Interactive file browser to select your database file
- **📊 Comprehensive Logging**: Complete activity logging with timestamps
- **🔄 Camera Refresh**: Click camera to refresh when it freezes after sleep mode
- **🛡️ Environment Validation**: Automatic creation of required folders and files
- **⚠️ Error Handling**: Enhanced error detection and user notifications
- **🔧 Better File Management**: Improved file handling and data validation

## Installation & Usage

### Option 1: Ready-to-use Executable (Recommended)

1. Download the latest executable file (`CCM Attend System.exe`) from the releases section
2. No additional installation required - the executable includes all necessary libraries and dependencies
3. **NEW**: The system will automatically create required folders on first run:
   - `day_data/` - For storing daily attendance records
   - `day_data/backup/` - For backup copies of daily records
   - `month_data/` - For monthly attendance summaries
   - `student_data/` - For student database
   - `logs/` - For system activity logs
4. **NEW**: Interactive database selection - choose your Excel database file on first run
5. Run the executable and start using the attendance system immediately

### Option 2: From Source (For Developers)

If you wish to modify the code or run from source:

1. Requirements:
   - Python 3.8+
   - Webcam or USB camera
   - Windows OS (tested on Windows 10/11)

2. Install the required packages:
   ```bash
   pip install opencv-python pandas pillow pyzbar icecream tkinter
   ```

3. Run the application:
   ```bash
   python improved_attendance_system.py
   ```

### Building Your Own Executable

To create a single `.exe` file:

```bash
# Simple build
pyinstaller "Attend system.spec"

# With icon and console (recommended for debugging)
pyinstaller --onefile --console --icon=ccmlogo_nKz_icon.ico --name="CCM Attendance System" improved_attendance_system.py

# Without console (for production)
pyinstaller --onefile --noconsole --icon=ccmlogo_nKz_icon.ico --name="CCM Attendance System" improved_attendance_system.py
```

## Using the System

### First Time Setup
1. **Launch**: Double-click the executable or run the Python script
2. **Database Selection**: Choose your Excel database file when prompted
3. **Camera Access**: Allow camera permissions when requested

### Daily Operations
1. **Sign-In Mode (Blue)**: Default mode for student check-ins
2. **QR Code Scanning**: Students scan their QR code IDs to check in
3. **Mode Switching**: Click "切换模式" (Switch Mode) to change to check-out mode (red)
4. **Sign-Out Mode (Red)**: Students scan again to check out
5. **Camera Issues**: Click anywhere on camera display to refresh if it freezes

### New Features Usage
- **📁 File Selection**: On first run, choose your database file through the file browser
- **🔄 Camera Refresh**: Simply click the camera display if it becomes unresponsive
- **📊 View Logs**: Check the `logs/` folder for detailed activity records
- **⚠️ Error Messages**: Clear error notifications guide you through any issues

## System Features in Detail

### Logging System
- **Daily Log Files**: `logs/attendance_system_YYYY_MM_DD.log`
- **Comprehensive Tracking**: All sign-ins, sign-outs, errors, and system events
- **Timestamped Entries**: Precise tracking of when events occur
- **Debug Information**: Helpful for troubleshooting issues

### File Management
- **Automatic Backup**: Every attendance record is automatically backed up
- **Environment Validation**: System creates missing folders automatically
- **Database Flexibility**: Choose any Excel database file with proper formatting
- **Data Integrity**: Built-in checks to prevent data corruption

### User Interface
- **Responsive Design**: Adapts to different screen sizes
- **Visual Indicators**: Clear color coding (Blue=Sign-in, Red=Sign-out)
- **Student Counter**: Shows current attendance vs total students
- **Status Updates**: Real-time feedback on all operations

## Data Structure

- **Daily Reports**: Individual Excel files for each day (`YYYY_MM_DD.xlsx`)
- **Monthly Reports**: Summarized attendance for each month (`Month YYYY.xlsx`)
- **Student Database**: Main student information file with attendance statistics
- **Activity Logs**: Detailed system logs (`logs/attendance_system_YYYY_MM_DD.log`)
- **Backup Files**: Timestamped backup copies in `day_data/backup/`

## Database Format Requirements

Your Excel database should include these columns:
- `编号` (Student ID) - Unique identifier for QR codes
- `孩子姓名(中)` (Student Name) - Student's name in Chinese
- `生日日期` (Birth Date) - For attendance rate calculations
- `attendance_days` - Total attendance days
- `attendance_by_month` - Monthly attendance count
- `attendance_rate` - Overall attendance rate
- `attendance_rate_by_month` - Monthly attendance rate

## Troubleshooting

### Common Issues
- **Camera not working**: Click the camera display to refresh
- **Database not found**: Use the file browser to select your Excel file
- **QR code not scanning**: Ensure good lighting and clear QR code
- **Student not found**: Check if student ID exists in database

### Log Files
Check the daily log files in `logs/` folder for detailed error information and system activity.

## Version History

### v2.0 (Latest)
- ✅ Interactive database file selection
- ✅ Comprehensive logging system
- ✅ Camera refresh functionality
- ✅ Automatic environment setup
- ✅ Enhanced error handling
- ✅ Improved user interface

### Previous Versions
- v1.6: Basic attendance tracking
- v1.5: Excel report generation
- v1.4: QR code integration

## License

[MIT License](LICENSE)

## Contact

For support or inquiries, please open an issue on this repository or contact the maintainer.