# MDU Result Portal Automation

## Description

This project automates the process of fetching student result data from the MDU result portal and saves it into a formatted Excel file. It uses Selenium WebDriver to interact with the website, extracts student details and marks, and organizes the data in an Excel sheet. The script also tracks processed students and supports resuming from the last processed row.

---

## Dependencies

- **Python 3.7+**
- [openpyxl](https://pypi.org/project/openpyxl/) (for Excel file handling)
- [selenium](https://pypi.org/project/selenium/) (for browser automation)
- [tkinter](https://docs.python.org/3/library/tkinter.html) (for message dialogs, included in standard Python)
- Microsoft Edge browser and [Edge WebDriver](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/) (ensure the driver matches your browser version)

**Install dependencies with:**
```sh
pip install openpyxl selenium
```

---

## How to Use This Program

1. **Clone or Download the Repository**
   ```sh
   git clone https://github.com/yourusername/your-repo-name.git
   cd your-repo-name/Automat-mdu-result-portal-2
   ```

2. **Prepare Your Data**
   - Edit the `main.py` file.
   - Update the `student_details` list with tuples of your students' registration numbers and roll numbers:
     ```python
     student_details = [
         (2312051612, 6071438),
         (2312051613, 6071439),
         # ...add more students...
     ]
     ```

3. **Ensure Edge WebDriver is Installed**
   - Download the correct version of Edge WebDriver for your Edge browser.
   - Place the WebDriver executable in your PATH or in the project directory.

4. **Run the Script**
   ```sh
   python main.py
   ```
   - The script will open the Edge browser, fetch results for each student, and save the data in `Student Result Data .xlsx`.
   - Progress and errors will be shown via message boxes.

5. **Check the Output**
   - The Excel file `Student Result Data .xlsx` will contain all fetched data, formatted and organized.
   - `Processed_data.txt` tracks which registration numbers have been processed.
   - `row count status .txt` tracks the current Excel row for resuming.

---

## Contribution Guidelines

We welcome contributions! Please follow these rules:

1. **Fork the repository** and create your branch from `main`.
2. **Write clear, concise commit messages**.
3. **Test your code** before submitting a pull request.
4. **Document any new features** or changes in the README.
5. **Respect the code style** and structure of the project.
6. **Open an issue** for major changes or feature requests before starting work.
7. **Be respectful and constructive** in all discussions.

---

**Thank you for your interest in improving this project!**
