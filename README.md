# Python Pizza Shop Project

A command-line pizza shop management system written in Python that handles customer orders, inventory, receipts, restocking, and basic financial tracking.

Originally developed for **CS 115**, the project demonstrates modular Python design, persistent data handling, CSV processing, Excel workbook integration, exception handling, and file-based application state.

## Highlights

- Processes customized pizza orders and calculates totals
- Tracks pizza-base and topping inventory
- Automatically updates inventory after purchases
- Generates customer receipt files
- Records sales and restocking expenses
- Displays profit-and-loss information and account balance
- Logs runtime errors for troubleshooting
- Separates application logic across multiple Python modules

## Technologies

- Python 3
- Pandas
- OpenPyXL
- xlwings
- CSV
- Microsoft Excel

## Project Structure

```text
Python-Pizza-Shop-Project/
│
├── Library/
│   ├── functions.py
│   └── log.py
├── prices/
│   ├── Base.csv
│   ├── Base_raw.csv
│   ├── Topping.csv
│   └── Topping_raw.csv
├── Receipts/
│   └── .gitkeep
├── .gitignore
├── pizza.py
├── pizza.xlsx
├── requirements.txt
└── README.md
```

`pizza.py` provides the command-line interface, while `Library/functions.py` contains the main ordering, inventory, restocking, and financial logic. Application data is stored using CSV files and an Excel workbook.

## Installation

Clone the repository:

```bash
git clone https://github.com/Ilkremer/Python-Pizza-Shop-Project.git
cd Python-Pizza-Shop-Project
```

Create and activate a virtual environment:

```bash
python -m venv .venv
```

Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

> Some financial functionality uses `xlwings` and requires Microsoft Excel, so the full application is primarily intended for Windows systems with Excel installed.

## Running the Program

```bash
python pizza.py
```

The program provides menu options for:

- Creating customer orders
- Viewing inventory
- Restocking ingredients
- Viewing profit-and-loss information
- Checking the current account balance

## Engineering Concepts Demonstrated

- Modular Python programming
- File I/O and persistent application data
- CSV parsing
- Excel workbook manipulation
- Input validation
- Exception handling and error logging
- Inventory-management logic
- Basic financial tracking
- Separation of application functionality across modules

## Current Limitations

This was originally an academic project rather than a production application. Current limitations include:

- Dependence on Excel and `xlwings`
- File-based storage instead of a database
- Command-line-only interface
- Limited automated testing
- Some application logic remains tightly coupled to workbook operations

## Potential Improvements

Future development could include:

- Replacing Excel storage with SQLite
- Adding automated tests and CI
- Separating business logic from the user interface
- Improving input validation
- Adding a graphical or web interface
- Adding reporting and analytics

## Author

**Isaac Kremer**  
Electrical and Computer Engineering Student

GitHub: [@Ilkremer](https://github.com/Ilkremer)