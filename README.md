# CCM Attendance System

A comprehensive QR code-based attendance tracking system designed for educational institutions. This system allows students to check in and out using QR codes, maintaining accurate attendance records on a daily and monthly basis.

![CCM Attendance System](CCMLogo.png)

## Features

- **QR Code Scanning**: Fast and efficient check-in/check-out with QR code scanning
- **Real-time Attendance Tracking**: View attendance status instantly
- **Daily and Monthly Reports**: Automatically generates Excel reports for day and month attendance
- **Student Database Integration**: Works with your existing student database
- **User-friendly Interface**: Clean, intuitive interface with clear visual indicators
- **Attendance Statistics**: Tracks attendance rates for each student
- **Backup System**: Automatic backup of attendance records

## Installation & Usage

### Option 1: Ready-to-use Executable (Recommended)

1. Download the latest executable file (`CCM Attend System.exe`) from the releases section
2. No additional installation required - the executable includes all necessary libraries and dependencies
3. Create the following folder structure in the same directory as the executable:
   - `day_data/` - For storing daily attendance records
   - `day_data/backup/` - For backup copies of daily records
   - `month_data/` - For monthly attendance summaries
   - `student_data/` - For student database
4. Place your student database Excel file in the root directory (named "CCM MASTER DATABASE_UPDATED YR 2024.xlsx" by default)
5. Run the executable and start using the attendance system immediately

### Option 2: From Source (For Developers)

If you wish to modify the code or run from source:

1. Requirements:
   - Python 3.8+
   - Webcam or USB camera
   - Windows OS (tested on Windows 10/11)

2. Install the required packages:
   ```
   pip install opencv-python pandas pillow pyzbar icecream
   ```

3. Run the application:
   ```
   python improved_attendance_system.py
   ```

## Using the System

1. Launch the application by double-clicking the executable or running the Python script
2. The system will start with the camera feed displayed on the left and attendance list on the right
3. Students scan their QR code IDs to check in (blue mode)
4. Click the "切换模式" (Switch Mode) button to change to check-out mode (red)
5. Students scan again to check out
6. Reports are automatically generated in the appropriate folders

## Data Structure

- **Daily Reports**: Individual Excel files for each day (`YYYY_MM_DD.xlsx`)
- **Monthly Reports**: Summarized attendance for each month (`Month YYYY.xlsx`)
- **Student Database**: Main student information file with attendance statistics

## License

[MIT License](LICENSE)

## Contact

For support or inquiries, please open an issue on this repository or contact the maintainer.