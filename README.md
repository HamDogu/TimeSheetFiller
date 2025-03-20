
# Timesheet Filler 📊

Welcome to the **Timesheet Filler Project**! This project is designed to scrape data from Outlook and perform various operations on the data, such as handling holidays and exporting timesheets. This README will guide you through the setup, usage, and features of the project.

## Project Structure 📁

Here’s a quick overview of the project files:

- **outlookScraper.py**: Script for scraping emails from Outlook.
- **holidays.py**: Script for managing holiday data.
- **main.py**: The main script that runs the project.
- **main.spec**: PyInstaller spec file for building the project.

## Features ✨

The project provides the following features:

- **Scrape Emails**: Extract emails from Outlook.
- **Parse Emails**: Extract relevant information from emails.
- **Handle Holidays**: Manage holiday data for a given year.
- **Export Timesheets**: Export timesheet data to various formats.

## Setup 🛠️

To get started with the project, follow these steps:

1. **Clone the Repository**:
   Use the following command to clone the repository:
   ```bash
   git clone https://github.com/yourusername/timesheet-export.git
   cd timesheet-export
   ```

2. **Install Dependencies**:
   Install the necessary Python libraries by running:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the Project**:
   To run the project, simply execute the main script:
   ```bash
   python main.py
   ```

## Usage 🚀

### Scraping Emails
The `outlookScraper.py` script contains functions to scrape data from Outlook. You can specify filters to extract only relevant emails.

### Handling Holidays
The `holidays.py` script provides functions to handle holidays. It can automatically exclude holidays when calculating timesheets.

### Main Execution
The `main.py` script is the entry point of the project. Running this script will execute the main function and trigger the various operations such as scraping, parsing, and exporting data.

## Building the Project 🏗️
To build the project into an executable, use the `main.spec` file with PyInstaller:

```bash
pyinstaller main.spec
```

This will generate the executable files in the `build` directory.

## Use Case 💼
Imagine you are a project manager who needs to generate timesheets for your team based on their email activity. This project allows you to:

- **Scrape Emails**: Automatically extract emails from Outlook.
- **Parse Emails**: Extract relevant information such as sender, recipient, and timestamps.
- **Handle Holidays**: Exclude holidays from the timesheet calculations.
- **Export Timesheets**: Generate timesheets in various formats for reporting and analysis.

## Contributing 🤝
Contributions are welcome! If you’d like to contribute to the project:

- Fork the repository.
- Create a new branch (`git checkout -b feature-branch`).
- Make your changes and commit them (`git commit -am 'Add new feature'`).
- Push to your fork (`git push origin feature-branch`).
- Create a pull request.

## License 📄
This project is licensed under the MIT License. See the LICENSE file for more details.

Feel free to update this README file with any additional information or changes to the project. Happy coding! 😊
