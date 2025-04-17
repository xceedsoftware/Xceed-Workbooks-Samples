![NuGet Downloads](https://img.shields.io/nuget/dt/Xceed.Workbooks.NET) ![Static Badge](https://img.shields.io/badge/.Net_Framework-4.0%2B-blue) ![Static Badge](https://img.shields.io/badge/.Net-5.0%2B-blue) [![Learn More](https://img.shields.io/badge/Learn-More-blue?style=flat&labelColor=gray)](https://xceed.com/en/our-products/product/workbooks-for-net) [![See Demo](https://img.shields.io/badge/Simple_Live_Demo-▶-brightgreen)](https://xceedsoftware.github.io/Xceed-Workbooks-Samples/) [![See Demo](https://img.shields.io/badge/Complex_Live_Demo-▶-red)](https://xceedsoftware.github.io/BlazorDocx-public/workbooksample)

[![Xceed Workbooks for .NET](./Resources/workbooks_header.png)](https://xceed.com/en/our-products/product/workbooks-for-net)


# Xceed Workbooks for .NET - Examples

This repository contains a variety of sample applications to help you get started with using the Xceed Workbooks for .NET in your own projects.

## Overview

Xceed Workbooks for .NET allows you to create or manipulate Microsoft Excel documents from your .NET applications, without requiring Excel or Office to be installed. It is fast, lightweight, and backed by a responsive support and development team. 

## About The Product

Xceed Workbooks for .NET provides an easy-to-use API to create or modify Excel .xlsx documents, offering complete control over content, formatting, and styling. It's ideal for generating reports and invoices programmatically, leveraging Excel’s familiar interface and rich editing features. It supports cell content modification, column and row resizing, formula calculation, data loading, and adding elements like pictures and hyperlinks. Key features include:

- **Intuitive and Simple API**: Designed for ease of use and efficiency.
- **Complete Control**: Modify content of cells, columns, and rows, create formatted tables, set and calculate formulas, load data, and add elements like pictures and hyperlinks.
- **Customization**: Style cells, rows, and columns using different fonts, alignments, and formatting settings.
- **Great for Reporting**: Use Excel documents as templates for reports and invoices.

For more information, please visit the [official product page](https://xceed.com/en/our-products/product/workbooks-for-net).

### Why Choose Xceed Workbooks for .NET?

- Developed by an experienced development team.
- Regular updates and new feature releases.
- Supports .NET 4.5, 5, 6, and 7.
- Comprehensive documentation and sample applications included.

## Getting Started with Xceed Workbooks for .NET.

To get started, clone this repository and explore the various sample projects provided. Each sample demonstrates different features and capabilities of Xceed Workbooks for .NET.

### Requirements
- Visual Studio 2015 or later
- .NET Framework 4.0 or later
- .NET 5.0 or later

### 1. Installing the Xceed Workbooks for .NET from nuget
To install the Xceed Workbooks for .NET from NuGet, follow these steps:

1. **Open your project in Visual Studio.**
2. **Open the NuGet Package Manager Console** by navigating to `Tools > NuGet Package Manager > Package Manager Console`.
3. **Run the following command:**
```sh
   dotnet add package Xceed.Workbooks.NET
```

4. Alternatively, you can use the NuGet Package Manager GUI:

1. Right-click on your project in the Solution Explorer.
2. Select Manage NuGet Packages.
3. Search for Xceed.Workbooks.NET and click Install.

![Nuget library](./Resources/nuget_sample.png)

### 2. Refering Xceed Workbooks for .NET library

1. **Add the reference with using statement at the top of the class**
   ```
   using Xceed.Workbooks.NET;
   ```
   
2. **Use the classes and elements from the namespace**
   ```c#
   using Xceed.Workbooks.NET;

   namespace BlazeDocX.Services
   {
       public class WorkBookCreator
       {
           private readonly IJSRuntime jsRuntime;
   
           public WorkBookCreator(IJSRuntime _jsRuntime)
           {
               jsRuntime = _jsRuntime;
           }
           public async Task CreateWorkbook(Accounting accounting)
           {
               using (var workbook = Workbook.Create("FileAsExcel"))
               {
                   ExportFile(accounting, workbook);
                   using MemoryStream memStream = new();
                   workbook.SaveAs(memStream);
                   await jsRuntime.InvokeVoidAsync("blazeDocX.downloadStream", memStream.GetBuffer(), $"FileAsExcel.xlsx");
               }
           }
       }
   }
   ```

   ### 3. How to License the Product Using the LicenseKey Property
To license the Xceed Forkbooks for .NET using the LicenseKey property, follow these steps:

1. **Obtain your license key** from Xceed. (Download the product from xceed.com or send us a request at support@xceed.com
2. **Set the LicenseKey property in your application startup code:**

   2.1 In case of WPF or Desktop app could be in the MainWindow
   ```csharp
   using System.Windows;

   public partial class MainWindow : Window
   {
       public MainWindow()
       {
           InitializeComponent();
           Xceed.Workbooks.NET.Licenser.LicenseKey = "Your-Key-Here";
       }
   }
   ```
   2.2 In case of ASP.NET application must be in Program.cs class
   ```csharp
   using System.Net;
   using System.Text.Json;
   using System.Text.Json.Serialization;
   ...
   using Xceed.Document.NET;
   ...
   Xceed.Workbooks.NET.Licenser.LicenseKey = "Your-Key-Here";
   ...
   var builder = WebAssemblyHostBuilder.CreateDefault(args);
   ```
4. Ensure the license key is set before any Workbooks class, instance or similar control is instantiated.

## Sample Applications
### Basic Usage
A simple example showing how to create a workbook with 3 worksheets with  into a specific cells with some style.

```csharp
    public static void AddWorksheets()
    {
      using( var workbook = Workbook.Create( SomeDirectory + @"AddWorksheets.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Add Worksheets";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill cells in 1st worksheet.
        worksheet.Cells[ "D5" ].Value = "This is the first Worksheet.";

        // Add a worksheet with default "SheetX" name.
        workbook.Worksheets.Add();

        // Fill cells in 2nd worksheet.
        workbook.Worksheets[ 1 ].Cells[ "D5" ].Value = "This is the second Worksheet.";

        // Add a worksheet with name "Third Sheet".
        workbook.Worksheets.Add( "Third Sheet" );

        // Fill cells in 3rd worksheet.
        workbook.Worksheets[ "Third Sheet" ].Cells[ "D5" ].Value = "This is the third Worksheet.";

        // Save workbook to disk.
        workbook.Save();
      }
    }

```

## Examples Overview

Below is a list of the examples available in this repository:

- **Annotations**: Demonstrates how to add and style a comment or note.
- **Cell**: Shows how to work with cells (sets the cell type as numeric, datetime, text, boolean... working with formulas, merging cells, mutiple fonts, replacing content, delete a range cell, insert a range cell).
- **Column**: Provides a number of examples about how works with columns (Cell access, Customize columns, Hide/unhide columns, Clear contents).
- **Hyperlink**: Demonstrates how to add hyperlink.
- **ImportData**: Shows how to use date from a outer source and display it (from arrays, collections, DataTables, csv).
- **Miscellaneous**: Demonstrates how to load data from a website.
- **Picture**: Provides examples how handle images (adding a picture, offset and shrink with offset).
- **Protection**: Shows how to use data protection (worksheet protection, worksheet with password, allow actions, remove protection, lock / unlock cells).
- **Row**: Demonstrates how to work and handle rows in worksheets (access, customize, hide/unhide).
- **SheetView**: Different ways and styles of sheets (freeze rows and columns, active cells, zoom and type view, split rows and columns).
- **Style**: All of a universe of methods about styles (text directions, text orientation, modify theme, alignments, borders, built in styles, fills, fonts, formatting, style on ranges).
- **Table**: Shows how add or remove tables.
- **Workbook**: Demonstrates working with workbooks (formulas, create a simple workbook, with filename, with stream, with url).
- **Worksheet**: Shows how to handle with worksheet (add new, calculate formulas, cell access, copy worksheets, insert, hide, move, remove, row access).

## Getting Started with the Samples

To get started with these examples, clone the repository and open the solution file in Visual Studio.

```bash
git clone https://github.com/your-repo/Xceed-Workbooks-Samples.git
cd xceed-workbooks-samples
```
Open the solution file in Visual Studio and build the project to restore the necessary NuGet packages.
  
## Documentation

For more information on how to use the Xceed Workbooks for .NET, please refer to the [official documentation](https://doc.xceed.com/xceed-workbooks-for-net/webframe.html#topic1.html).

## Licensing

To receive a license key, visit [xceed.com](https://xceed.com) and download the product, or contact us directly at [support@xceed.com](mailto:support@xceed.com) and we will provide you with a trial key.

## Contact

If you have any questions, feel free to open an issue or contact us at [support@xceed.com](mailto:support@xceed.com).

---

© 2024 Xceed Software Inc. All rights reserved.
