
## Overview

This Google Apps Script project follows a **Client-Server Architecture** typical of web-based Add-ons, using Google Sheet as the host environment and Apps Script as the backend service layer. The primary goal is to perform a data-intensive audit using performance-optimized Google services.

## Component Breakdown

### 1. Client (Frontend - `Sidebar.html`)

* **Technology:** HTML, CSS, and plain JavaScript.
* **Role:** Provides the user interface (UI) and handles user interaction (button clicks).
* **Communication:** Uses the asynchronous `google.script.run` API to call server-side functions.
* **Progress Polling:** The `pollProgress` function uses `setInterval` to periodically call `getProgress()` on the server, ensuring the UI's progress bar and file count (`#bar`, `#fileCount`) are updated in near real-time without blocking the scan.

### 2. Server (Backend - `Code.gs`)

* **Technology:** Google Apps Script (JavaScript/V8 runtime).
* **Role:** Contains all the business logic, handles interactions with Google Drive, Google Sheets, and manages state.
* **Key Functions:**
    * **`onOpen()`:** Creates the custom menu item in the Google Sheet.
    * **`scanEntireDrive()`:** The main audit loop.
        * Uses **`Drive.Files.list`** (Advanced Service) for paginated, efficient fetching of file metadata.
        * Calls `getPermissionsInBatch` to perform parallel permission checks.
        * Accumulates file data into the `data` array and saves it to `ScriptProperties` at each page turn.
    * **`getPermissionsInBatch()`:** The performance bottleneck solver. It uses **`UrlFetchApp.fetchAll`** to concurrently make multiple Drive API calls (`/permissions` endpoint) for a batch of files, significantly reducing latency compared to serial API calls.
    * **`writeToSheet()`:** Handles the final step: reading the complete data from `ScriptProperties`, structuring the sheet, and writing the final results.
    * **`reScanAsync()`/`startBackgroundScanTrigger()`:** Manages time-based triggers for background or delayed execution.

### 3. Data and State Management

* **`PropertiesService` (ScriptProperties):** Used to persist the application state between iterative function calls and client polls.
    * `PROP_KEY`: Stores the **cumulative JSON array** of audited file data.
    * `PAGE_TOKEN_KEY`: Stores the **next page token** from the Drive API list response, crucial for resuming the scan from where it left off.
    * `TOTAL_COUNT_KEY`: Stores the **total number of files** calculated on the first run, preventing repeated slow counting.
    * `PROCESSED_COUNT`/`progress`: Used by the client's `pollProgress` function to update the UI.

### 4. External Services

* **`Drive API` (Advanced Service):** Provides the `Drive.Files.list` method for high-performance file searching and pagination.
* **`UrlFetchApp`:** Used within `getPermissionsInBatch` to make concurrent HTTP requests to the Drive API (specifically, the `/permissions` endpoint) for performance.
* **Google Sheet:** The persistent data storage and UI host environment.