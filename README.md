# FacultySubmissionToOutlook25Live
```markdown

## Overview

🚀 **Excited to Share My Latest Project!** 🚀

As **Lab Assistant Supervisor** at Joliet Junior College, I developed a robust Python automation framework that integrates our **25Live** scheduling system with **Microsoft Outlook**. This solution streamlines the process of scheduling lab exams by reading CSV/Excel files, processing relevant test data, and automatically populating events into both 25Live and Outlook calendars.

**Key Benefits:**
- **Streamlines Lab Scheduling**: Filters/validates CSV/Excel data and calculates accurate end times with configurable buffer periods.
- **Enhances Efficiency**: Automates keystroke input (via `pyautogui`) to populate Outlook and 25Live without manual intervention.
- **Reduces Manual Errors**: Eliminates double-entry, minimizes human error, and ensures data accuracy.

It’s been incredibly rewarding to see how leveraging automation has simplified complex workflows at JJC, saving time and improving accuracy. I’m excited to continue pushing the boundaries of tech-driven innovation!

## Features
- **Data Filtering & Validation**  
  Automatically removes irrelevant entries (e.g., “laptop” rows), checks for existing 25Live or Outlook entries, and ensures consistent data formats.

- **Automated Date & Time Calculations**  
  Parses input time strings, converts them to 12-hour or 24-hour formats, and adds buffer periods (e.g., 15 minutes) to create accurate end times.

- **Flexible File Handling**  
  Supports both `.csv` and `.xlsx` file types, automatically detecting worksheets with season-year indicators (e.g., `SP24`).

- **Browser & Keystroke Automation**  
  Uses `webbrowser` to open Outlook/25Live and `pyautogui` for simulating keystrokes and clicks to populate scheduling fields.

- **Error Handling & Logging**  
  Prints descriptive error messages for unsupported file formats or missing data, aiding in debugging.

## Tech Stack
- **Python 3.9+** (or higher)
- **pandas** for data manipulation
- **openpyxl** for `.xlsx` file support
- **pyautogui** for automated keystroke and mouse interactions
- **re**, **datetime**, **csv** for parsing and date/time operations
- **webbrowser** for opening Outlook/25Live links

## Installation

1. **Clone the Repository**  
   ```bash
   git clone https://github.com/sngalamou/FacultySubmissionToOutlook25Live.git
   cd FacultySubmissionToOutlook25Live
   ```

2. **Create & Activate a Virtual Environment** (optional, but recommended)  
   ```bash
   python -m venv venv
   source venv/bin/activate     # On macOS/Linux
   venv\Scripts\activate        # On Windows
   ```

3. **Install Dependencies**  
   ```bash
   pip install -r requirements.txt
   ```
   Make sure `requirements.txt` includes `pandas`, `openpyxl`, and any other relevant packages.

## Usage

1. **Prepare Your Data File**  
   - Place your `.csv` or `.xlsx` file in the same directory as the script.
   - Ensure it contains the correct headers (date, time, location, etc.).

2. **Run the Script**  
   ```bash
   python FacultSubs_to_25Live.py
   ```
   - The script will prompt for **semester season** (`SP` or `FL`) and **year** (last two digits).
   - It will then parse the data, filter out invalid rows, and prompt you to insert events into Outlook/25Live.

3. **Automation in Action**  
   - Outlook Calendar and 25Live tabs will open automatically in your default browser.
   - `pyautogui` will simulate keystrokes and mouse actions to populate event details.

4. **Check Your Schedules**  
   - After execution, verify events in Outlook Calendar and 25Live to confirm successful insertion.
   - The script logs any skipped events (e.g., duplicates or past-dated events).

## Project Structure
```
.
├── FacultSubs_to_25Live.py      # Main automation script
├── requirements.txt             # Python dependencies
├── README.md                    # Project documentation
```

## Contributing
1. **Fork** the repository
2. **Create** your feature branch: `git checkout -b feature/my-awesome-feature`
3. **Commit** your changes: `git commit -m "Add awesome feature"`
4. **Push** to the branch: `git push origin feature/my-awesome-feature`
5. **Open** a Pull Request on GitHub

## License
This project is not licensed for public use. Please do not use or adapt the code.

## Acknowledgments
- **Supervisor:** I would like to thank my supervisor, Paul Schroeder, for his support, guidance, and the autonomy he granted me, which were essential to the successful completion of this project.
- **Joliet Junior College** for providing a real-world environment to develop and test the automation script.
- **The Python Community** for numerous libraries (`pandas`, `pyautogui`, `openpyxl`) that make development faster and easier.
- **You** for checking out this project!

---

**Check out the source code here:**  
[Github Repo](https://github.com/sngalamou/FacultySubmissionToOutlook25Live)

**Follow me on LinkedIn:**  
[LinkedIn Profile](https://www.linkedin.com/in/sngalamou/)
