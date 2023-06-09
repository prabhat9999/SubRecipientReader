# Subrecipient Reader

The Subrecipient Reader is a C# console application that reads data from Excel spreadsheets containing subrecipient information. Specifically, it reads the "G. Other Direct Costs" section of each spreadsheet and extracts the subrecipient name and subaward amount.

## Dependencies
The Subrecipient Reader uses the following dependencies:

#### ExcelDataReader
#### ExcelDataReader.DataSet

## Usage

To use the Subrecipient Reader, follow these steps:

Clone this repository to your local machine
Open the solution file SubRecipient.sln in Visual Studio
Build the solution to restore the NuGet packages and compile the application
Place your Excel spreadsheets containing subrecipient information in the files folder in the root directory of the application
Run the application from within Visual Studio or from the command line using dotnet SubrecipientReader.dll
The application will output the subrecipient data to the console, including the file name, subrecipient name, and subaward amount. It will also output the total subaward amount for each subrecipient.
Note: The application currently only supports reading .xlsx files.

## Contributing

If you find a bug or have a feature request, please create an issue in the GitHub repository. Pull requests are welcome!

## License

This code is released under the MIT License.
