# ObjectToCsv

## C# Library: Object List to CSV Converter
[![ObjectToCsv](https://img.shields.io/nuget/v/ObjectToCsv.svg?style=plastic&logo=nuget)](https://www.nuget.org/packages/ObjectToCsv)
[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat)](https://github.com/Open-range-org/ObjectToCsv/blob/main/LICENSE)

## Overview
This library provides an easy way to convert a list of objects into a CSV (Comma-Separated Values) file. It's particularly useful when you need to export data from your application to a format that can be easily imported into spreadsheet software or other tools.

## Features
- Converts a list of objects (such as a collection of custom classes) into a CSV file.
- Handles escaping special characters (like commas and double quotes) within the data.
- Supports custom delimiters (e.g., semicolon, tab) if needed.

## Installation
1. **Using NuGet Package Manager:**
   - Open your project in Visual Studio.
   - Go to **Tools > NuGet Package Manager > Manage NuGet Packages for Solution**.
   - Search for "ObjectToCsv" and install it.
   - ObjectToCsv is a popular library that simplifies reading and writing CSV files in C#.

2. **Manual Installation:**
   - Download the ObjectToCsv library from NuGet.
   - Add the reference to your project.
## Custom Attributes
1. In C#, when working with models, it’s common to encounter scenarios where the property names in the model differ from the column names in a CSV file. To address this issue, we can use the `Header` attribute. This attribute facilitates the mapping between CSV columns and model properties.
2. Additionally, there’s another attribute called `Ordinal`. The purpose of this attribute is to help the library identify the order of CSV headers, which can be useful during data processing.

## Usage Example
Suppose you have a class called `Person` with properties like `Id`, `Name`, `Age`, and `Email`. You want to convert a list of `Person` objects to a CSV file.

1. Create a list of `Person` objects:
2. Add `Header` and `Ordinal` attribute for properties.

```csharp
public class PersonObject
{
    [Header("Id")]
    [Ordinal(0)]
    public Guid Id { get; set; }

    [Header("Email")]
    [Ordinal(1)]
    public string Email { get; set; }
	
    [Header("User Name")]
    [Ordinal(2)]
    public string Name { get; set; }

    [Header("Age")]
    [Ordinal(3)]
    public string Age { get; set; }

    [Header("Email")]
    [Ordinal(4)]
    public string Email { get; set; }
}
```

```csharp
using ObjectToCsv;


// Retrieve a list of PersonObject instances to be converted to CSV.
List<PersonObject> users = GetTestObjectList();

// Define a custom date format for DateTime fields in the CSV output.
string dateFormat = "dddd, dd MMMM yyyy"; // Example: "Monday, 01 January 2024"

// Use the CsvUtil library to convert the list of PersonObjects into a CSV-formatted string.
// The dateFormat parameter is optional; if not provided, the default format "dd MMMM yyyy" will be used.
var csvString = CsvUtil.BindCsv<PersonObject>(users, dateFormat);

// csvString now contains the CSV representation of the users list.
```


## Contributing
Feel free to contribute to this library by submitting pull requests or reporting issues on the GitHub repository.

## License
This library is released under the MIT License. See the LICENSE file for details.