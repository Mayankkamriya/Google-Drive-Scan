# Google Drive Scan Script

This project is a Google Apps Script solution designed to **audit file metadata and sharing permissions** across a user's Google Drive. It operates within a Google Sheet, using a custom sidebar for a user-friendly interface. It leverages the **Google Drive API Advanced Service** for efficiency and parallel processing to handle large numbers of files without hitting the Apps Script execution time limits.

## ✨ Features

* **Interactive Sidebar:** A dedicated UI (`Sidebar.html`) within the Google Sheet for controlling the audit process.
* **Efficient Scanning:** Uses the Drive API's `Files.list` method for pagination and faster data retrieval.
* **Parallel Permission Check:** Employs `UrlFetchApp.fetchAll` to batch and parallelize fetching file permission counts, significantly speeding up the most intensive part of the audit.
* **Asynchronous Processing:** The scan runs in the background, updating progress in the UI using a polling mechanism.
* **Detailed Output:** Exports key file metadata and permission counts to a dedicated sheet tab (`Drive Audit`).
* **Audit Limit:** Configurable limit (`MAX_FILES`) to prevent excessively long scans on huge drives.

## 🛠️ Setup and Installation

### 1. Create the Project

1.  Open a new Google Sheet.
2.  Go to **Extensions** > **Apps Script**.
3.  Rename the default file to `Code.gs` and paste the provided Google Apps Script code into it.
4.  Click the **+** icon next to "Files" and select **HTML**. Name the file `Sidebar.html` and paste the provided HTML code into it.

### 2. Enable Advanced Service

For the script to access the `Drive.Files.list` and permission endpoints, the Drive API must be explicitly enabled.

1.  In the Apps Script editor, click the **Services** icon ($\text{+}$ sign in the left sidebar).
2.  Scroll down and select **Drive API**.
3.  Click **Add**.

### 3. Run the Setup Function

1.  In the Apps Script editor, select the **`onOpen`** function from the dropdown menu (next to the Run button).
2.  Click the **Run** button.
3.  **Authorization:** The first time, you will be prompted to grant permissions. Accept the permissions.
4.  **Reload** the Google Sheet. A new custom menu item, **Soluvery**, should appear.

## 🏃 Usage

1.  In the Google Sheet, navigate to **Soluvery** > **Open Drive Audit** to launch the sidebar.
2.  **Scan My Drive:** Click this button to start the audit. The scan runs iteratively in the background and updates the progress bar.
3.  **Write to Sheet:** Once the scan reports completion (100% progress), click this button to take the collected data from the script's memory and write it to the **Drive Audit** sheet tab.
4.  **Re-scan (Async):** This resets the scan progress, deletes the old temporary data, schedules a new background scan, and automatically starts polling the progress.