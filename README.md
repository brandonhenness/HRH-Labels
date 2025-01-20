# MEDITECH Labels

A Python script designed to generate custom labels based on data extracted from Excel files. This tool is tailored for use in a MEDITECH-driven inventory or transfer management workflow.

## Features

- Extracts data from an Excel file and processes it for label generation.
- Generates PDF labels with a configurable HTML template and CSS styling.
- Labels include dynamic symbols and colors based on predefined inventory and transfer rules.
- Processes batches of up to 30 labels per page for efficiency.

## Requirements

- Python 3.8 or newer
- Required Python packages:
  - `openpyxl` for reading Excel files
  - `blabel` for generating labels
  - `tkinter` for the file selection dialog (included in the Python standard library)

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/brandonhenness/MEDITECH-Labels.git
   cd MEDITECH-Labels
   ```

2. Install the required dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Ensure the following additional files are available in the same directory as the script:
   - `label_template.html` (HTML template for the labels)
   - `style.css` (Stylesheet for label customization)

## Usage

1. Run the script:

   ```bash
   python meditech_labels.py
   ```

2. A file selection dialog will appear. Choose an Excel file with the following columns:
   - **Inventory** (Column A)
   - **Item ID** (Column B)
   - **Description** (Column C)
   - **Transfer Status** (Column P)

3. The script will process the data and generate a PDF (`qrcode_and_label.pdf`) with the labels.

### Label Symbols and Colors

- **Red Diamond (◆):** Non-transferable items not in `MORDEP`.
- **Blue Diamond Outline (◇):** Transferable items in `MORDEP`.
- **Purple Bullseye (◈):** Non-transferable items in `MORDEP`.
- **No Symbol:** All other cases.

## Example

Given an Excel file with inventory data:

| Inventory | Item ID | Description        | Transfer |
|-----------|---------|--------------------|----------|
| MORDEP    | 12345   | Example Item 1     | Y        |
| OTHER     | 54321   | Example Item 2     | N        |

The script generates labels as per the above rules, styling, and layout.

## Contributing

Contributions are welcome! If you'd like to improve this project, please submit a pull request or open an issue on GitHub.

## License

This project is licensed under the [GPL-3.0 License](LICENSE).

