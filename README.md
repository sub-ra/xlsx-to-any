# XLSX to Various Formats Converter

This Python tool converts Excel spreadsheets (`.xlsx` files) into multiple formats, including Markdown, HTML, CSV, JSON, and others, facilitating the transfer of Excel tables into various environments. It's designed to be an open-source project, where contributions for improvements and extensions are welcome.

## Features

- Converts each Excel worksheet into a variety of formats: Markdown, HTML, CSV, JSON, YAML, XML, Plain Text, Jira, MediaWiki.
- User-friendly GUI interface with a dropdown menu to select the desired output format.
- Supports toggling options to include or exclude hidden rows and columns in the conversion.

## Prerequisites

Before running this tool, ensure you have the following libraries installed:

```bash
pip install -r requirements.txt
```

This command installs the following libraries:

- `openpyxl` for reading Excel files.
- `tabulate` for generating tables in different formats.
- `PyYAML` for generating YAML files.
- `pandas` for handling data.

Please note that `tkinter`, which is used for the GUI interface, is part of the Python standard library and might not be included by default on some systems. If you're using a system where `tkinter` isn't included by default, you'll need to install it using your system's package manager.

## Installation

To set up this tool, clone the repository or download the source code:

```bash
git clone https://github.com/yourusername/xlsx-to-mdtable.git
cd xlsx-to-mdtable
```

## Usage

Follow these steps to convert your Excel files:

1. Run the script.
2. Select the Excel file you want to convert.
3. Choose the worksheets you want to convert to Markdown.
4. Toggle the option to include or exclude hidden rows and columns.

## Supported Output Formats
- md - Markdown
- html - HTML
- csv - CSV
- json - JSON
- yaml - YAML
- xml - XML
- txt - Plain Text
- jira - Jira Markdown
- mediawiki - MediaWiki Markup

## Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Contact

Project Link: [https://github.com/sub-ra/xlsx-to-any](https://github.com/sub-ra/xlsx-to-any)

## Acknowledgements

- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [tabulate](https://pypi.org/project/tabulate/)
- [tkinter](https://docs.python.org/3/library/tkinter.html)
- [PyYAML](https://pyyaml.org/)
- [pandas](https://pandas.pydata.org/)
