# College Attendance Management System

This is a desktop application for managing student attendance, built with Python's Tkinter library. It provides a graphical user interface (GUI) to add students, mark their attendance, and notify parents via email in case of absence.

## Features

- **Student Management**: Add new students with their ID, name, and parent's email. You can also delete existing students.
- **Attendance Tracking**: Mark students as "Present" or "Absent" for any selected date.
- **Date Selection**: An interactive calendar to easily pick the date for attendance marking.
- **Real-time Updates**: The student list shows updated attendance percentages.
- **Data Storage**: All student information and attendance records are stored in an Excel file (`attendance.xlsx`).
- **Email Notifications**: Send an email to a student's parent to inform them about their child's absence and their current attendance percentage.

## Requirements

- Python 3
- The following Python libraries:
  - `tkinter` (usually included with Python)
  - `tkcalendar`
  - `openpyxl`

## Installation

1.  **Clone the repository or download the source code.**

2.  **Install the required libraries:**
    ```bash
    pip install tkcalendar openpyxl
    ```

## How to Run the Application

1.  **Make sure you have the following files in the same directory:**
    - `pro.py` (the main script)
    - `Amrita.png` (for the background image, or you can change the path in the script)

2.  **Run the script from your terminal:**
    ```bash
    python pro.py
    ```

## How to Use

1.  **Add a Student**:
    - Enter the Student ID, Name, and Parent's Email in the respective fields.
    - Click "Add Student". The student will be added to the list and saved to `attendance.xlsx`.

2.  **Mark Attendance**:
    - Select a date using the "Select Date" button.
    - Select a student from the list.
    - Click "Mark Present" or "Mark Absent".
    - The attendance for that student on the selected date will be recorded in the Excel file.
    - The selection will automatically move to the next student in the list to allow for quick marking.

3.  **Send Email**:
    - Select a student from the list.
    - Click "Send Email to parent". An email will be sent to the registered parent's email address if their child was marked absent.
    - **Note**: You need to configure your email credentials in `pro.py` for this to work. Find the following lines and replace them with your details:
      ```python
      sender_email = "your_email@gmail.com"
      sender_password = "your_app_password" 
      ```
      It is recommended to use an "App Password" for Gmail if you have 2-Factor Authentication enabled.

4.  **View Report**:
    - Click "View Attendance Report" to open the `attendance.xlsx` file.

5.  **Delete a Student**:
    - Select a student from the list.
    - Click "Delete Student". You will be asked for confirmation before the student is removed.

## ✍️ Author

Made with ❤️ by **Lalith kumar raju Somalaraju**.

Connect with me:

*   [<img src="https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white" />](https://github.com/lalith-kumar-raju)
*   [<img src="https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white" />](https://www.linkedin.com/in/lalith-kumar-raju-somalaraju/)
*   [<img src="https://img.shields.io/badge/Instagram-E4405F?style=for-the-badge&logo=instagram&logoColor=white" />](https://www.instagram.com/_lalith_kumar_raju_/)
*   [<img src="https://img.shields.io/badge/Email-D14836?style=for-the-badge&logo=gmail&logoColor=white" />](mailto:ssivaprasadraju1978@gmail.com)

--- 