# XLSX to Various Formats Converter

This Python tool converts Excel spreadsheets (`.xlsx` files) into multiple formats, including Markdown, HTML, CSV, JSON, and others, facilitating the transfer of Excel tables into various environments. It's designed to be an open-source project, where contributions for improvements and extensions are welcome.

![GUI](https://github.com/sub-ra/xlsx-to-any/assets/87712870/8f417c9c-d10e-4941-b3ae-4d4541b1e378)

## Features

- Converts each Excel worksheet into a variety of formats: Markdown, HTML, CSV, JSON, YAML, XML, Plain Text, Jira, MediaWiki.
- User-friendly GUI interface with a dropdown menu to select the desired output format.
- Supports toggling options to include or exclude hidden rows and columns in the conversion.

## Prerequisites

All Prerequisites are in the pyproject.toml
`Python
Openpyxl
tabulate
pyyaml
pandas
tk`

## Installation

To clone the repository or download the source code:

```bash
git clone https://github.com/yourusername/xlsx-to-any.git
cd xlsx-to-any
```

To create a Wheel package from the project, you can use Poetry. Install Poetry and navigate to the project directory via terminal, then run:
```bash
poetry build
```
After that, you can install the Wheel Package:
```bash
pip install path/to/whl/package/NAMEOFTHEPACKAGE.whl
```

## Usage

Follow these steps to convert your Excel files:

1. Run the script with `xlsx_to_any`
2. Select the Excel file you want to convert.
3. Choose the worksheets you want to convert.
4. Toggle the option to include or exclude hidden rows and columns.
5. Select the format.

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
