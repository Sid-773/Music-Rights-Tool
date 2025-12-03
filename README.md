# Excel Processing Web Application

This is a simple web application that allows you to upload an Excel file, process its data according to specific rules, and export the results as individual Excel files or a single zip file.

## How to Share This Application Using GitHub Pages

GitHub Pages is a free service that lets you host a website directly from a GitHub repository. Follow these steps to make your application shareable with a single link.

### Step 1: Create a GitHub Account

If you don't already have one, sign up for a free account at [https://github.com/join](https://github.com/join).

### Step 2: Create a New Repository

1.  Once you are logged in, click the **+** icon in the top-right corner and select **New repository**.
2.  Give your repository a name (e.g., `excel-data-tool`).
3.  Make sure the repository is set to **Public**. This is required for GitHub Pages to work on the free plan.
4.  You can skip adding a README, .gitignore, or license for now.
5.  Click **Create repository**.

### Step 3: Upload Your Project Files

1.  On your new repository's main page, click the **Add file** button and select **Upload files**.
2.  Drag and drop the three files from this project (`index.html`, `style.css`, and `script.js`) into the upload area.
3.  Scroll down and click the **Commit changes** button.

### Step 4: Enable GitHub Pages

1.  In your repository, click on the **Settings** tab (located on the top navigation bar).
2.  In the left-hand sidebar, scroll down and click on **Pages**.
3.  Under the "Build and deployment" section, for the "Source", select **Deploy from a branch**.
4.  Under the "Branch" section, make sure the selected branch is `main` (or `master`) and the folder is set to `/ (root)`.
5.  Click **Save**.

### Step 5: Access and Share Your Live Site

1.  After saving, GitHub will start deploying your site. This may take a few minutes.
2.  Once it's ready, a green bar will appear at the top of the Pages settings screen with a link to your live site. The URL will look something like this:
    `https://<your-username>.github.io/<your-repository-name>/`
3.  You can now share this single link with anyone. When they open it, they will be able to use the application directly in their web browser without needing to download any files.

That's it! Your application is now live on the web.

---

## How the Code Works (Step-by-Step)

This section explains the step-by-step execution flow of the application, which can be helpful for debugging issues related to Excel file templates.

1.  **File Selection**: You select an Excel file via the drag-and-drop zone or the "Choose File" button. This triggers the `handleFile` function.

2.  **File Loading**:
    *   The "Submit" button is disabled and a loading indicator appears.
    *   The browser's `FileReader` API reads the file into memory.

3.  **Excel Parsing (`onload` event)**:
    *   The powerful **SheetJS (XLSX)** library parses the entire Excel workbook.
    *   The script loops through **every sheet** in the workbook.
    *   It checks if the sheet name meets one of two conditions:
        *   The name is **"FILM"** (case-insensitive).
        *   The name matches the pattern **"Episode #NNN"** (e.g., "Episode #201").
    *   If a sheet is valid, the script stores its data and dynamically creates a **"Distribution ID" input field** on the page, labeled with that sheet's name.
    *   Once all sheets are checked, the "Submit" button is re-enabled and the loader disappears.

4.  **Data Processing (on "Submit" click)**:
    *   The `displayAllEpisodeData` function is called.
    *   It loops through each valid sheet's stored data and runs the `processSheetData` function, which performs the following key actions:
        *   It reads the **Distribution ID** (from the corresponding input) and the **Start Row** you entered.
        *   It iterates through the rows of the sheet, starting from your specified row.
        *   **Column Splitting**: It finds the text **"Written by"** (case-insensitively) in the first column. The text before becomes the `SONG_TITLE`, and the text after becomes the `WRITERS`.
        *   **Artists Column**: It takes the data from the second column as the `ARTISTS`. If the value is "N/A", it is converted to a blank string.
        *   **Duplicate Removal**: It checks for and removes any duplicate rows based on the combination of `SONG_TITLE`, `ARTISTS`, and `WRITERS`.
    *   The final, clean data is used to generate an HTML table for each valid sheet on the webpage.

5.  **Exporting**:
    *   When you click "Export All as ZIP", the script re-processes the data for all displayed sheets, generates an `.xlsx` file for each one in memory, and bundles them into a single `.zip` file for download.

### Common Excel Template Issues to Check

If a file isn't working, it is often due to "template anomalies." Check your Excel file for the following:

*   **Valid Sheet Name**: The script now relies on the sheet's *name*. Make sure the tabs you want to process are named either "FILM" or follow the "Episode #NNN" convention.
*   **Merged Cells**: The parser expects a simple grid. Merged cells, especially in the first few columns, can disrupt data processing.
*   **Hidden Characters**: Data copied from other sources can contain invisible characters that may interfere with text matching (like the "Written by" search).
*   **Unexpected Formatting**: While less common, complex cell formatting or embedded objects can sometimes cause issues for the parser.