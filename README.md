# Financial Analysis and Simulation Project of Warren Buffet Strategy

## Overview

This repository contains a Python script and an Excel file for conducting financial analysis and simulating the performance of a stock investment strategy. The project combines data analysis, API integration, and historical simulations to evaluate the effectiveness of a specific investment strategy.

## Table of Contents

- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
- [Project Components](#project-components)
- [Python Script](#python-script)
- [Excel Simulation](#excel-simulation)
- [Usage](#usage)
- [Results](#results)
- [Contributing](#contributing)
- [License](#license)

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
