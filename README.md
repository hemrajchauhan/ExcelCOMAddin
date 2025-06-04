# ExcelCOMAddin

**ExcelCOMAddin** is a COM-based add-in for Microsoft Excel, developed using C# and .NET Framework. It extends Excel's functionality by introducing custom commands and automations, enhancing productivity for users who require specialized tools within their spreadsheets.

## Features

- **Directory Tools**: 
  - File Browser
  - Delete Empty Folders
- **File Operations**: 
  - Create, Delete, Move, Copy, Rename, Replace
  - Unhide or Delete multiple files efficiently
- **Text Tools**: 
  - Convert text to Upper, Lower, Title Case
  - Trim or Join strings
  - Extract Left/Mid/Right characters
- **Excel Utilities**:
  - Import data quickly
  - Combine and Split worksheets
  - Generate File Lists
- **Hyperlink Management**: 
  - Quick hyperlink insertion

## UI Preview
Here’s how the custom ribbon appears in Excel:
![ExcelCOMAddin Ribbon UI](https://github.com/user-attachments/assets/34dd8686-3e4c-4b32-9cf7-113aa48e51d3)

## Getting Started

### Prerequisites

- Microsoft Excel (2016 or later)
- Visual Studio 2019 or newer
- .NET Framework 4.7.2 or higher

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/hemrajchauhan/ExcelCOMAddin.git

2. Open the MiscAddins.sln solution file in Visual Studio.

3. Build the solution to generate the COM Add-in DLL.

4. Register the COM Add-in:
   ```bash
   regsvr32 path\to\your\add-in.dll

5. Open Excel to verify that the add-in has been successfully integrated.

### Usage
Once installed, the add-in will add a new tab or group in the Excel ribbon. Users can access the custom functionalities provided by the add-in through these new UI elements.

### Contributing
Contributions are welcome! Please fork the repository and submit a pull request with your enhancements.

### License
This project is licensed under the GNU General Public License v3.0.
