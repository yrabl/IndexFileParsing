# Ada XML to Excel Converter

Ada XML to Excel Converter is a Windows Forms application that processes XML files and converts them into an Excel file. The application allows users to specify the data path and Excel file, and provides options to delete or rename the XML files after processing.

## Features

- Browse and select the data path containing XML files.
- Browse and select the output Excel file.
- Option to delete XML files after processing.
- Option to rename XML files based on their metadata.
- Log messages to track the processing status.
- Save and load settings from a configuration file.
- **Multi-language support** (English and Hebrew).

## Requirements

- .NET 8
- Visual Studio 2022
- OfficeOpenXml library

## Installation

1. Download the latest release from GitHub: [Latest Release](https://github.com/yrabl/IndexFileParsing/releases/latest)

2. Extract the `Installer.zip` file.

3. Run `Installer.msi` to install the application.

4. Follow the on-screen instructions to complete the installation.

## Usage

1. Run the application.

2. Use the "Browse" button to select the data path containing the XML files.

3. Use the "Browse" button to select the output Excel file.

4. Check the "Delete XML files after processing" option if you want to delete the XML files after processing.

5. Check the "Rename Files" option if you want to rename the XML files based on their metadata.

6. Select the application language from the dropdown (English or Hebrew).

7. Click the "Process" button to start processing the XML files and generate the Excel file.

8. The log messages will be displayed in the log text box to track the processing status.

## Configuration

The application saves and loads settings from a configuration file named `config.json`, which is stored in `C:\ProgramData\AdaXmlToExcel\`. The configuration file contains the following settings:

- `DataPath`: The path to the data files.
- `ExcelFile`: The path to the Excel file.
- `DeleteFiles`: A value indicating whether to delete files after processing.
- `RenameFiles`: A value indicating whether to rename files after processing.


The application ensures that the configuration directory and file exist before attempting to read or write settings. If the file does not exist, it is created with default values:

```json
{
  "DataPath": "",
  "ExcelFile": "",
  "DeleteFiles": false,
  "RenameFiles": false
}
```

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue to discuss any changes or improvements.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgements

- [OfficeOpenXml](https://github.com/EPPlusSoftware/EPPlus) for providing the library to work with Excel files.
