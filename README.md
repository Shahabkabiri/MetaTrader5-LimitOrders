# MetaTrader 5 Limit Order Execution

This Python script connects to MetaTrader 5 using the MetaTrader 5 API and executes limit orders based on data from an Excel file. The script reads orders from the Excel file, determines the order type, and executes pending limit orders in MetaTrader 5.

## Prerequisites

- MetaTrader 5 (MT5) platform installed
- MetaTrader 5 API (MetaTrader 5 Gateway) configured for Python
- Python
- MetaTrader 5 Python package (`MetaTrader5`)

## Usage

1. Clone this repository or download the script.
2. Install the required Python packages using `pip install MetaTrader5 pandas`.
3. Modify the path to your orders Excel file in the script (`OrderFilePathAndName`).
4. Customize the order execution logic as needed (e.g., adjusting order parameters).
5. Run the script using `python script_name.py`, where `script_name.py` is the name of your script.

## Description

This script performs the following tasks:

- Initializes the MetaTrader 5 API connection.
- Reads limit orders from an Excel file.
- Filters orders that are marked as used and selects the last 100 orders.
- Executes buy and sell limit orders based on the order type and specified price levels.

The script is designed to automate the execution of pending limit orders in the MetaTrader 5 platform.

## Customization

You can customize the script by modifying the order execution logic, adjusting order parameters, or adding additional functionality based on your trading strategy.

## License

This project is licensed under the [MIT License](LICENSE).
