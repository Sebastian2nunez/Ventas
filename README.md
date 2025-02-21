Ventas

A Python-based tool for processing and analyzing sales and cost data from Excel files. This application helps you clean, map, and compute various metrics (such as unit and total costs, margins, and quantities) based on multiple document types (e.g., electronic invoices, credit notes, export invoices, etc.). It leverages a graphical interface to allow users to select input files and define starting rows and filters for processing the data.

Features

Data Cleaning:
Cleans and formats columns (e.g., SKU codes) for consistent processing.

Cost Calculation:
Computes unit and total costs by grouping data by SKU, with special handling for document types (e.g., making certain costs negative for credit notes or electronic receipts).

Dynamic Mapping:
Creates mapping dictionaries to link SKUs with their corresponding cost values from different data sources.

Flexible Filtering:
Supports filtering data with customizable conditions and creating pivot tables for further analysis.

User-Friendly Interface:
Uses Tkinter to provide file dialogs for selecting Excel files and to input parameters like starting row numbers and filters.


Requirements

Python: 3.8 or later

Libraries:

pandas

openpyxl

Tkinter (usually included with standard Python installations)



Additional packages may be required for HTML table parsing (such as lxml or BeautifulSoup4) if your costs file is in HTML format.

Installation

1. Clone the repository:

git clone https://github.com/Sebastian2nunez/Ventas.git
cd Ventas


2. Create and activate a virtual environment (optional but recommended):

python -m venv venv
source venv/bin/activate  # On Windows, use `venv\Scripts\activate`


3. Install the required dependencies:

pip install pandas openpyxl

If you need HTML parsing support, you might also install:

pip install lxml beautifulsoup4



Usage

1. Run the application:

python ventas3.py


2. Select the input files:

A sales file (Excel format, e.g., .xlsx)

A costs file (Excel or HTML format; note that HTML files are processed using pd.read_html)

An annual file (if available)



3. Configure the parameters:

Specify the starting row for reading the sales and annual files.

Optionally set filters for pivot table creation.

The GUI will provide prompts and messages for any errors or necessary input.



4. Execute the processing:

Once all files and parameters are set, click on the button to execute the program.

The tool will read, process, and calculate the required metrics (like “Costo Neto Unitario”, “Costo Total Neto”, “Total venta”, etc.), handling various document types appropriately.



5. Output:

Processed data can be further manipulated or saved into new Excel files, depending on how you integrate additional logic within the script.




Customization

Functions Overview:
The script is modular, with functions dedicated to:

Cleaning DataFrames (limpiar_filas)

Calculating totals (calcular_costos_totales)

Mapping SKUs to cost values (crear_diccionarios_mapeo)

Filtering and creating pivot tables (filtrar_y_crear_tabla_dinamica)


Feel free to modify or extend these functions to suit your specific data processing needs.

GUI Enhancements:
The current version uses Tkinter for basic file selection and parameter entry. You can expand the GUI for a more robust user experience if required.


Contributing

Contributions are welcome! If you have suggestions, improvements, or bug fixes, please feel free to fork the repository and submit a pull request.
