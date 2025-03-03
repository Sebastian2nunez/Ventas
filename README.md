# Sales

A Python tool to process and analyze sales data from Excel files.  
This project uses **pandas** for data manipulation, **openpyxl** for working with Excel, and **tkinter** for a graphical interface that facilitates file selection and parameter configuration.

## Description

The main script, `ventas3.py`, implements several functions that allow you to:

- **Extract and Clean Data:**  
  Functions such as `limpiar_filas` and `obtener_tipo_y_numero` handle formatting and extracting relevant information from DataFrame columns.

- **Cost Calculation and Assignment:**  
  It calculates total and unit costs (using functions like `calcular_costos_totales`, `Costos_negativos`, and `asignar_costos`) for sales documents and credit notes.

- **Dynamic Table Generation and Filtering:**  
  With the `filtrar_y_crear_tabla_dinamica` function, pivot tables are created to summarize information, while other functions allow dynamic filtering based on conditions.

- **Graphical User Interface (GUI):**  
  Utilizes **tkinter** to enable the user to select sales, cost, and annual data files, enter the starting row for data reading, and configure filters for generating pivot tables.

The script reads, processes, and transforms the data to subsequently assign costs and generate summaries that facilitate the analysis of sales information.

## Requirements

- **Python 3.7+** (it is recommended to use a virtual environment)  
- **Libraries:**
  - [pandas](https://pandas.pydata.org/)
  - [openpyxl](https://openpyxl.readthedocs.io/)
  - **tkinter** (included with most Python installations)  
- Other standard modules such as `re`, `os`, and `importlib`.

## Installation

1. **Clone the repository:**

   ```bash
   git clone https://github.com/Sebastian2nunez/Ventas.git
   cd Ventas
Create a virtual environment (optional but recommended):

python -m venv venv
source venv/bin/activate  # On Linux/Mac
venv\Scripts\activate     # On Windows

# Usage
python ventas3.py
A graphical window will open allowing you to:

Select the sales file (Excel format).
Select the costs file (Excel or HTML format, as per the code).
Select the annual file (optional).
Enter the starting row number for data reading.
Configure filters for generating pivot tables.
The script will perform the following operations:

Read and preprocess the files.
Clean columns and format data.
Calculate unit and total net costs, adjusting values based on document types (e.g., converting certain values to negative for receipts and credit notes).
Create summaries and pivot tables to facilitate data analysis.
Project Structure
ventas3.py: Main script containing all functions for data manipulation, analysis, and visualization.
README.md: This documentation file.
Key Features
Data Processing:
Cleaning, conversion, and formatting of key columns (such as SKU and costs).

Advanced Calculations:
Calculation of total cost per SKU, assignment of unit and total costs, and margin calculation.

Graphical User Interface (GUI):
Use of tkinter for an interactive user experience in file selection and parameter configuration.

Flexibility:
Allows dynamic filtering for pivot table creation and analysis of various document types (receipts, credit notes, invoices, etc.).

Contributing
If you wish to contribute to this project:

Fork the repository.
Create a branch for your new feature or bug fix.
Make your changes and submit a pull request.
Any contribution is welcome!

License
[Specify the license here, e.g., MIT License]

Contact
For inquiries, suggestions, or reporting bugs, please contact through the GitHub repository or send an email to your contact address.
