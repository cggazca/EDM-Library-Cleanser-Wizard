# Project Overview

This project is a Python-based GUI application named "EDM Library Wizard". It is built using PyQt5 and is designed to assist users in converting data from Microsoft Access databases (`.mdb` or `.accdb`) into Excel files. The application then facilitates the generation of XML files formatted for the "EDM Library Creator v1.7.000.0130".

The wizard guides users through a multi-step process:
1.  **Data Source Selection**: Users can either convert an Access database to an Excel file or use an existing Excel file.
2.  **Column Mapping**: Users can map columns from their Excel sheets to specific fields required for the XML output (e.g., Manufacturer, Manufacturer Part Number).
3.  **Data Normalization**: The application provides tools to clean and normalize the data, including a side-by-side comparison of original and modified data.
4.  **XML Generation**: Finally, the application generates XML files based on the mapped and cleaned data.

The project uses several key libraries, including:
-   `pandas` for data manipulation.
-   `sqlalchemy` for database interaction.
-   `pyodbc` to connect to Access databases.
-   `xlsxwriter` for creating Excel files.
-   `PyQt5` for the graphical user interface.
-   `fuzzywuzzy` for string matching.

## Building and Running

### Dependencies

To run the application, you need to install the required Python packages. The project includes `requirements_wizard.txt` and `requirements_test.txt` files, which list the dependencies.

**Install dependencies using pip:**
```bash
pip install -r requirements_wizard.txt
```

**Windows-Specific Requirement:**
For handling Access databases, the Microsoft Access Database Engine (ODBC Driver) must be installed. It can be downloaded from [here](https://www.microsoft.com/en-us/download/details.aspx?id=54920).

### Running the Application

Once the dependencies are installed, you can run the application from the command line:

```bash
python edm_wizard.py
```

### Creating an Executable

The project also includes instructions for creating a standalone executable file on Windows using `PyInstaller`.

1.  **Install PyInstaller:**
    ```bash
    pip install pyinstaller
    ```

2.  **Build the executable:**
    ```bash
    pyinstaller --onefile --windowed --name "EDM_Library_Wizard" edm_wizard.py
    ```
    The executable will be created in the `dist/` folder.

## Development Conventions

-   The project is structured with a main `edm_wizard.py` file that serves as the entry point.
-   The UI is organized into pages, with each page corresponding to a step in the wizard. These pages are located in the `edm_wizard/ui/pages` directory.
-   The application logic is separated from the UI, with data processing and other utility functions located in the `edm_wizard/utils` directory.
-   The application uses threads for long-running tasks to keep the UI responsive.
-   The code includes error handling and provides informative messages to the user.
-   The project includes a `.gitignore` file to exclude common Python and development-related files from version control.
