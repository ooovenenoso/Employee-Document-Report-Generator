# Employee Document Report Generator

This repository contains a **Google Apps Script** project that automates the process of generating and sending employee document status reports. The script monitors a Google Sheet, creates a formatted Google Doc report, sends it via email, and logs all actions. It also organizes generated reports in a designated Google Drive folder and maintains a history of sent reports within cell notes.

## Features

- **Automated Trigger**: Uses an installable trigger to monitor a specific column in a Google Sheet. When a designated value (other than `(En blanco)`) is selected, the report generation process starts automatically.
- **Custom Report Generation**: Generates a Google Doc report containing employee data and document statuses, including:
  - Employee details (Site, Name, Position)
  - Selected report type (e.g., "All Documents", "Certificate of Health", etc.)
  - Document expiration information
  - Report generation date
- **Email Notification**: Sends a beautifully formatted email (with the company logo and custom styles) to the employee's email address. If no email is provided, a default email is used.
- **Google Drive Integration**: Automatically moves the generated Google Doc into a specified folder (created if it doesn’t exist) in your Google Drive.
- **Logging and History**:
  - Logs every action in a dedicated "Logs" sheet.
  - Appends a history entry (with date and time) in the note of the cell that triggered the report.
- **User Feedback**: Includes a link in the email for users to report any issues.

## Installation

1. **Create a Google Sheet** with the required structure. The project expects a sheet named `"Actualizado"` for employee data and a trigger column (Column U).
2. **Open the Apps Script Editor** from your Google Sheet:
   - Go to **Extensions > Apps Script**.
3. **Copy and Paste** the entire script from this repository into the script editor.
4. **Run the Permissions Function**:
   - Execute the function `requestAllPermissions()` to prompt for the necessary permissions (Gmail, Drive, and Sheets).
5. **Set Up the Trigger**:
   - Run the function `createInstallableTriggers()` to create the installable trigger for handling edits.
6. **Customize (Optional)**:
   - Update any column indices or sheet names if your setup differs.
   - Modify email content or report styling as needed.

## Usage

- **Triggering a Report**:  
  When you change the value in column U (for rows beyond the header) in the `"Actualizado"` sheet to any value other than `(En blanco)`, the script will:
  - Generate a report based on the selected criteria.
  - Send an email with the report attached.
  - Log the action and update the note with a timestamp.
  - Reset the triggering cell to `(En blanco)` after sending.
  
- **Reporting Issues**:  
  Users can click on the "Report a Problem" link in the email to provide feedback or report issues via a Google Form.

## Contributing

Contributions, issues, and feature requests are welcome. Feel free to check the [Issues](https://github.com/) section for open topics or open a new issue to start a discussion.

## License

This project is provided **as-is** without any warranty. Feel free to use and modify the code for your own purposes.  

## Disclaimer

This project uses generic data and configurations for demonstration purposes. Ensure you customize the project to meet your organization's security and data handling standards before deployment.

---

Happy coding!
