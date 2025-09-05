# Excel Dynamic Viewer

A modern WPF application built with C# that dynamically reads and displays Excel files with varying column structures. This solution uses free and open-source libraries to handle Excel files without any licensing restrictions.

![Excel Dynamic Viewer](https://img.shields.io/badge/Excel-Viewer-green.svg)
![WPF](https://img.shields.io/badge/WPF-MVVM-blue.svg)
![License](https://img.shields.io/badge/License-MIT-lightgrey.svg)
![.NET](https://img.shields.io/badge/.NET-9.0-purple.svg)

## Features

- **Dynamic Excel Handling**: Automatically detects columns and data types from any Excel file
- **Multiple Format Support**: Reads .XLS, .XLSX, and CSV files
- **Modern UI**: Clean, professional interface with Material Design principles
- **MVVM Architecture**: Proper separation of concerns for maintainability
- **Free & Open Source**: Uses ExcelDataReader (MIT License) - no licensing fees

## Screenshot

- will update it soon!

## Supported File Formats

- ✅ Microsoft Excel (.xlsx)
- ✅ Microsoft Excel 97-2003 (.xls) 
- ✅ Comma Separated Values (.csv)

## Technology Stack

- **.NET Framework**: .NET 9+
- **UI Framework**: WPF (Windows Presentation Foundation)
- **Architecture**: MVVM (Model-View-ViewModel)
- **Excel Library**: ExcelDataReader (MIT License)
- **Dependency Injection**: Built-in MVVM pattern

## Installation

### Prerequisites

- [.NET 9+](https://dotnet.microsoft.com/download)
- Windows OS (Windows 10/11 recommended)
- Visual Studio 2022 or Visual Studio Code

### Packages

- ExcelDataReader(3.7.0*)
- ExcelDataReader.DataSet(3.7.0*)
- CommunityToolkit.Mvvm(8.4.0*)
- Microsoft.Xaml.Behaviors.Wpf(1.1.135*)

### Clone the Repository

```bash
git clone https://github.com/kathan555/ExcelDynamicViewer.git
cd ExcelDynamicViewer
