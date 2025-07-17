# AEP XLSX Comparator

A simple Python application with a Tkinter GUI to compare Adobe Experience Platform debugger exports for production and development environments, filtering out unwanted rows and highlighting any differences.

## Features

- **Clean**: Removes specified attributes (from a configurable list) that are not required for comparison.
- **Compare**: Generates separate **Production** and **Development** sheets.
- **Highlight**: Creates a **Differences** sheet, highlighting any changed cells in yellow.
- **Configure**: Editable `config.json` to control which rows are removed.
- **GUI**: User-friendly Tkinter interface for selecting files, monitoring progress, and viewing completion status.

## Prerequisites

- Python 3.7 or above
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)

```bash
pip install pandas openpyxl
```

## Installation

1. Clone the repository:
   ```bash
   ```

git clone [https://github.com/](https://github.com/)\<your‑username>/aep-xlsx-comparator.git cd aep-xlsx-comparator

````
2. Install the dependencies (see **Prerequisites**).

## Usage

1. Run the application:
   ```bash
python app.py
````

2. On first run, a `config.json` file will be generated in the project root. Review and adjust the list of rows to remove as needed.
3. In the GUI:
   - Click **Browse** to select the **Production** XLSX export.
   - Click **Browse** to select the **Development** XLSX export.
   - Click **Go** to begin processing.
4. On completion, an `comparison_output.xlsx` file is created with three sheets:
   - `Production`
   - `Development`
   - `Differences` (with changed cells highlighted)

## Configuration

The `config.json` file contains a JSON array of row names to exclude from both files. To modify:

```json
[
    "Timestamp",
    "Time Since Page Load",
    "Initiator",
    "frame",
    "hitId",
    "isMultiSuiteTagging",
    "isTruncated",
    "reportSuiteIds",
    "returnType",
    "trackingServer",
    "version",
    ".a",
    ".activitymap",
    ".c",
    "a.",
    "Activity Map Link",
    "Activity Map Page",
    "Activity Map Page Type",
    "Activity Map Region",
    "activitymap.",
    "Audience Manager Blob",
    "Audience Manager Location Hint",
    "Browser Window Height",
    "Browser Window Width",
    "c.getPreviousValue",
    "c.getQueryParam",
    "c.pt",
    "Character Set",
    "ClickMap Object ID",
    "ClickMap Object Tag Name",
    "ClickMap Page ID",
    "ClickMap Page ID Type",
    "Color quality",
    "Context Data",
    "Cookie Domain",
    "Cookies Enabled",
    "Currency Code"
]
```

You may add or remove entries as required.

## Licence

This project is licensed under the MIT Licence. See the [LICENSE](LICENSE) file for details.

## Contributing

1. Fork the repository.
2. Create a feature branch:
   ```bash
   ```

git checkout -b feature/my‑feature

````
3. Commit your changes:
   ```bash
git commit -m "Add my feature"
````

4. Push to the branch:
   ```bash
   ```

git push origin feature/my‑feature

```
5. Open a pull request.

Please ensure code is well‑documented and adheres to PEP 8 style guidelines.

## Acknowledgements

- Based on the Adobe Experience Platform debugger export format.
- Utilises `pandas`, `openpyxl` and `tkinter`.

```
