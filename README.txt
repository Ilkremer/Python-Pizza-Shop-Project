# Python Pizza Shop Project

A command-line pizza shop management system written in Python. The program simulates the basic operations of a pizza shop, including processing customer orders, tracking inventory, restocking ingredients, generating receipts, and maintaining financial records.

This project was originally created for **CS 115** and demonstrates Python programming concepts such as functions, loops, exception handling, file I/O, CSV processing, and Excel workbook manipulation.

## Features

* Create customized pizza orders
* Select pizza sizes and toppings
* Automatically calculate order totals
* Generate customer receipt files
* Track pizza base and topping inventory
* Warn when inventory is low or empty
* Individually restock ingredients
* Perform a full inventory restock
* Track sales and restocking expenses
* View sales and expenses
* Calculate profit and loss
* Display the pizza shop's current account balance
* Log program errors to an error log

## Technologies Used

* **Python 3**
* **Pandas** — reads and processes pricing data
* **OpenPyXL** — reads and modifies Excel inventory and financial data
* **xlwings** — recalculates formulas in the Excel workbook
* **CSV** — stores pizza size and topping pricing
* **Microsoft Excel** — stores inventory and financial information

## Project Structure

```text
Python-Pizza-Shop-Project/
│
├── pizza.py
├── pizza.xlsx
├── ErrorLog.txt
│
├── Library/
│   ├── functions.py
│   └── log.py
│
├── prices/
│   ├── Base.csv
│   ├── Base_raw.csv
│   ├── Topping.csv
│   └── Topping_raw.csv
│
└── Receipts/
```

### Main Files

**`pizza.py`**
The main entry point for the program. It displays the pizza shop menu and allows the user to access the different features of the system.

**`Library/functions.py`**
Contains the main functions used for ordering pizzas, modifying inventory, restocking products, tracking finances, and displaying financial information.

**`Library/log.py`**
Contains functionality used to record errors encountered while the program is running.

**`pizza.xlsx`**
Stores the pizza shop's inventory and profit-and-loss information.

**`prices/`**
Contains CSV files containing pizza size, topping, and restocking price information.

**`Receipts/`**
Stores text receipts generated when customers place orders.

**`ErrorLog.txt`**
Stores information about exceptions encountered during program execution.

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/Ilkremer/Python-Pizza-Shop-Project.git
cd Python-Pizza-Shop-Project
```

### 2. Install the Required Python Packages

```bash
pip install pandas openpyxl xlwings
```

Because the current program uses `xlwings` to recalculate formulas in `pizza.xlsx`, it is best run on a system with **Microsoft Excel installed**.

## Running the Program

From the root directory of the repository, run:

```bash
python pizza.py
```

Depending on your system, you may instead need:

```bash
python3 pizza.py
```

The main menu will display:

```text
Please select a menu item.
1. Order a Pizza
2. View inventory levels
3. Restock inventory (for shop managers)
4. View the Profit and Loss statement
5. View the current bank balance.
6. Exit the program
```

Enter the number corresponding to the desired action.

## Ordering a Pizza

When placing an order, the program:

1. Requests the customer's name and phone number.
2. Displays the available pizza sizes.
3. Allows the customer to select a pizza size.
4. Displays the available toppings.
5. Allows multiple toppings to be selected.
6. Calculates the total price.
7. Creates a receipt in the `Receipts/` directory.
8. Deducts the selected ingredients from inventory.
9. Records the sale in the Excel workbook.

Receipt filenames contain the customer's name, date, and time.

Example:

```text
John_09-09-2026_14.30.15.txt
```

## Inventory Management

Inventory is stored in `pizza.xlsx`.

When an order is placed, the appropriate pizza base and toppings are automatically deducted from inventory.

The program also warns the user when an item is:

* Running low
* Completely out of stock

The restocking system provides two options:

### Individual Restock

Restocks selected ingredients back to their full inventory level.

### Full Restock

Restocks all ingredients below full inventory.

The cost of restocking is recorded as an expense in the profit-and-loss statement.

## Financial Tracking

The program tracks both sales and expenses using the `P and L` worksheet in `pizza.xlsx`.

Users can view:

* Individual sales
* Restocking expenses
* Total sales
* Total expenses
* Overall profit or loss
* Current account balance

The original shop balance is initialized at **$1,000.00**.

## Error Handling

Major program operations are wrapped in exception handling. Errors encountered during execution are written to:

```text
ErrorLog.txt
```

This provides a record that can be used to diagnose problems without terminating the entire application.

## Concepts Demonstrated

This project demonstrates several fundamental programming and software development concepts:

* Modular Python programming
* Functions and reusable code
* Loops and conditional logic
* User input validation
* Exception handling
* Reading CSV files
* Reading and modifying Excel workbooks
* File creation and management
* Inventory management logic
* Basic financial tracking
* Persistent application data
* Error logging

## Potential Future Improvements

Possible improvements to the project include:

* Replace the Excel workbook with SQLite
* Remove the dependency on Microsoft Excel and `xlwings`
* Add a graphical user interface
* Add automated tests
* Improve input validation
* Prevent orders for out-of-stock ingredients
* Add multiple-pizza orders
* Add quantities for toppings and inventory
* Add employee and manager authentication
* Convert the application into a web-based ordering system
* Add reporting and sales analytics

## Author

**Isaac Kremer**

Electrical and Computer Engineering Student

GitHub: [@Ilkremer](https://github.com/Ilkremer)
