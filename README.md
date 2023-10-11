# Financial Analysis and Simulation Project of Warren Buffet Strategy

## Overview

This repository contains a Python script and an Excel file for conducting financial analysis and simulating the performance of a stock investment strategy. The project combines data analysis, API integration, and historical simulations to evaluate the effectiveness of a specific investment strategy.

## Table of Contents

- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
- [Project Components](#project-components)
- [Python Script](#python-script)
- [Google Sheet Simulation](#google-simulation)
- [Google Sheet Analysis](#google-analysis)
- [Usage](#usage)
- [Results](#results)
- [Contributing](#contributing)

## Getting Started

### Prerequisites

Before using this project, ensure that you have the following prerequisites in place:

- Python 3.x
- Required Python packages (specified in the script)
- Excel software to open and use the simulation file
- API Key (more details in [Python Script](#python-script))

## Project Components

The project includes the following components:

1. **Python Script**: A Python script that retrieves financial data from an API, performs financial analysis, and calculates a score for selected stocks.

2. **Excel Simulation**: An Excel file that simulates historical stock investments based on the Python script's analysis. It allows you to explore different investment scenarios and evaluate the strategy's performance.

## Python Script

The Python script performs the following tasks:

- Retrieves financial data for a list of stock symbols from the financialmodelingprep API.
- Analyzes the data to calculate a score for each stock.
- Saves the results, including company information, scores, and recommendations, in the 'Analysis.txt' file.
- Provides a customizable and data-driven approach to stock selection.

### Usage

1. Clone this repository to your local machine.

2. Install the required Python packages by running the following command:

   ```bash
   pip install pandas numpy openpyxl
   ```

3.  Replace the API key in the Python script:

    In the analyze function, replace 'ce82b6a14287d6b24fdcaf5468401b12' with your own API key.

4. Run the Python script:

   ```bash
   python WBStrategy.py
   ```

5. The script will generate an 'Stocks.xlsx' file with the analysis results. ( **It take many hours to analyse all 35 000 stocks** )

### Results

The 'Stocks.xlsx' file generated by the Python script contains the analysis results, including company information, scores, and recommendations. For each year and for each stock you will get a score. More a stock gets close to 100%, more satisfies Warren Buffet criteria to invest in a stock.

## Google Sheet Simulation

The Excel file provides a simulation of historical stock investments based on the analysis performed by the Python script.

### Usage

1. Open a copy of this GoogleSheet : https://docs.google.com/spreadsheets/d/1BTwxfBKV1M9La5yX6W8ae-mf6Dwf9p2ABG2SrWBDa_s/edit?usp=sharing .

2. You can explore each socket's to know how this Google Sheets work but the main Tabs are the "portfolio" one, you can see the evolution of the portfolio from 01/01/2002 to 08/08/2023. You can change on the top right (in green) when you want to enter/exit from a stock. More those one get close to 100%, less stocks there will be in your portfolio.

## Google Sheet Analysis

1. Open a copy of this GoogleSheet : https://docs.google.com/spreadsheets/d/1IQvaiyNn9g66IUcMnfh851EniZu_vCtUYAIqxXgLssg/edit?usp=sharing .

2. You can explore each socket's to know how this Google Sheets work but the main Tabs are the "Sumary" to see results and "Data" to put your Portoflio data in the collumns "With inflation". It will compare your porfolio with the cac40, STOXX Europe 600, SnP 500 and ETF world.

## Contributing

If you would like to contribute to this project, please open an issue or submit a pull request. Contributions and improvements are welcome!
