/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – ImportData Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to import data to a worksheet when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;

namespace Xceed.Workbooks.NET.Examples
{
  public class ImportDataSample
  {
    #region Private Members

    private const string ImportDataSampleResourcesDirectory = Program.SampleDirectory + @"ImportData\Resources\";
    private const string ImportDataSampleOutputDirectory = Program.SampleDirectory + @"ImportData\Output\";

    #endregion

    #region Constructors

    static ImportDataSample()
    {
      if( !Directory.Exists( ImportDataSample.ImportDataSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ImportDataSample.ImportDataSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void ImportArrays()
    {
      using( var workbook = Workbook.Create( ImportDataSample.ImportDataSampleOutputDirectory + @"ImportArrays.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Import Arrays";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        worksheet.Cells[ "B4" ].Value = "Import a vertical array of strings:";
        worksheet.Cells[ "B4" ].Style.Font.Color = Color.Blue;

        // Define a string array, the import options(vertical by default) and call the ImportData function.
        var stringData = new string[] { "First", "Second", "Third", "Fourth" };
        var stringImportOptions = new ImportOptions() { DestinationTopLeftAddress = "B5" };
        worksheet.ImportData( stringData, stringImportOptions );


        worksheet.Cells[ "G4" ].Value = "Import an horizontal array of DateTimes:";
        worksheet.Cells[ "G4" ].Style.Font.Color = Color.Blue;

        // Define a DateTime array, the import options and call the ImportData function.
        var dateTimeData = new DateTime[] { new DateTime( 2022, 10, 10 ), new DateTime( 2020, 1, 15 ), new DateTime( 2021, 10, 11 ) };
        var dateTimeImportOptions = new ImportOptions() { DestinationRowId = 4, DestinationColumnId = 6, IsLinearDataVertical = false };
        worksheet.ImportData( dateTimeData, dateTimeImportOptions );


        worksheet.Cells[ "B14" ].Value = "Import a vertical array of Players:";
        worksheet.Cells[ "B14" ].Style.Font.Color = Color.Blue;

        // Define a user object array, the import options(vertical by default, show propertyNames) and call the ImportData function.
        var userObjectData = new Player[] 
        { 
          new Player() { Name = "Tom Sawyer", Team = Team.Miami_Ducks, Number = 9 },
          new Player() { Name = "Mike Smith", Team = Team.Chicago_Hornets, Number = 18 },
          new Player() { Name = "Kelly Tomson", Team = Team.LosAngelese_Raiders, Number = 33 },
          new Player() { Name = "John Graham", Team = Team.NewYork_Bucs, Number = 7 },
        };
        var userObjectImportOptions = new UserObjectImportOptions() { DestinationTopLeftAddress = "B15", IsPropertyNamesShown = true };
        worksheet.ImportData( userObjectData, userObjectImportOptions );


        worksheet.Cells[ "G14" ].Value = "Import a 2D array of doubles:";
        worksheet.Cells[ "G14" ].Style.Font.Color = Color.Blue;

        // Define a 2D array of doubles, the import options(vertical by default, show propertyNames) and call the ImportData function.
        var twoDData = new double[,] { { 11d, 22d, 33d, 44d }, { 55d, 66d, 77d, 88d }, { 99d, 100d, 101d, 102d } };
        var twoDImportOptions = new ImportOptions() { DestinationTopLeftAddress = "G15"};
        worksheet.ImportData( twoDData, twoDImportOptions );

        // AutoFit columns 1 to 3(columnId starts at 0), starting at row 14 up to row 19 (rowId starts at 0), with column sizes from 0 to 255.
        worksheet.Columns[ 1, 3 ].AutoFit( 0, 255, 14, 19 );

        // AutoFit columns G to I, starting at row 4 up to row 5 (rowId starts at 0), with column sizes from 0 to 255.
        worksheet.Columns[ "G", "I" ].AutoFit( 0, 255, 4, 5 );

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: ImportArrays.xslx\n" );
      }
    }

    public static void ImportCollections()
    {
      using( var workbook = Workbook.Create( ImportDataSample.ImportDataSampleOutputDirectory + @"ImportCollections.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Import Collections";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        worksheet.Cells[ "B4" ].Value = "Import a vertical list of strings:";
        worksheet.Cells[ "B4" ].Style.Font.Color = Color.Blue;

        // Define a list of strings, the import options(vertical by default) and call the ImportData function.
        var stringData = new List<string>() { "First", "Second", "Third", "Fourth" };
        var stringImportOptions = new ImportOptions() { DestinationTopLeftAddress = "B5" };
        worksheet.ImportData( stringData, stringImportOptions );


        worksheet.Cells[ "H4" ].Value = "Import a vertical List of Players:";
        worksheet.Cells[ "H4" ].Style.Font.Color = Color.Blue;

        // Define a list of user objects, the import options(vertical by default, specify PropertyNames and show propertyNames) and call the ImportData function.
        var userObjectData = new List<Player>()
        {
          new Player() { Name = "Tom Sawyer", Team = Team.Miami_Ducks, Number = 9 },
          new Player() { Name = "Mike Smith", Team = Team.Chicago_Hornets, Number = 18 },
          new Player() { Name = "Kelly Tomson", Team = Team.LosAngelese_Raiders, Number = 33 },
          new Player() { Name = "John Graham", Team = Team.NewYork_Bucs, Number = 7 },
        };
        var userObjectImportOptions = new UserObjectImportOptions() { DestinationTopLeftAddress = "H5", PropertyNames = new string[] { "Name", "Team" }, IsPropertyNamesShown = true };
        worksheet.ImportData( userObjectData, userObjectImportOptions );


        worksheet.Cells[ 13, 1 ].Value = "Import an horizontal ObservableCollection of int:";
        worksheet.Cells[ 13, 1 ].Style.Font.Color = Color.Blue;

        // Define a list of ints, the import options and call the ImportData function.
        var intData = new List<int>() { 1, 2, 3, 4, 5 };
        var intImportOptions = new ImportOptions() { DestinationRowId = 14, DestinationColumnId = 1, IsLinearDataVertical = false };
        worksheet.ImportData( intData, intImportOptions );


        worksheet.Cells[ "M4" ].Value = "Import a vertical Dictionary of Players:";
        worksheet.Cells[ "M4" ].Style.Font.Color = Color.Blue;

        // Define a Dictionary of Players, the import options and call the ImportData function.
        var dictionaryData = new Dictionary<int, Player>()
        {
          { 1, new Player() { Name = "Tom Sawyer", Team = Team.Miami_Ducks, Number = 9 } },
          { 2, new Player() { Name = "Mike Smith", Team = Team.Chicago_Hornets, Number = 18 } },
          { 3, new Player() { Name = "Kelly Tomson", Team = Team.LosAngelese_Raiders, Number = 33 } },
          { 4, new Player() { Name = "John Graham", Team = Team.NewYork_Bucs, Number = 7 } },
        };
        var dictionaryImportOptions = new UserObjectImportOptions() { DestinationTopLeftAddress = "M5" };
        worksheet.ImportData( dictionaryData, dictionaryImportOptions );

        // AutoFit all columns, starting at row 4 up to row 10 (rowId starts at 0), with column sizes from 0 to 255.
        worksheet.Columns.AutoFit( 0, 255, 4, 10 );

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: ImportCollections.xslx\n" );
      }
    }

    public static void ImportDataTables()
    {
      using( var workbook = Workbook.Create( ImportDataSample.ImportDataSampleOutputDirectory + @"ImportDataTables.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Import DataTables";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        worksheet.Cells[ "B4" ].Value = "Import a DataTable and style it:";
        worksheet.Cells[ "B4" ].Style.Font.Color = Color.Blue;

        // Define a dataTable, the import options(show specific ColumnNames) and call the ImportData function.
        var dataTable = new DataTable( "Employees" );
        dataTable.Columns.Add( "Name", typeof( string ) );
        dataTable.Columns.Add( "Position", typeof( string ) );
        dataTable.Columns.Add( "Experience", typeof( double ) );
        dataTable.Columns.Add( "Salary", typeof( int ) );
        dataTable.Rows.Add( "Jenny Melchuck", "Project Manager", 11.5d, 77000 );
        dataTable.Rows.Add( "Cindy Gartner", "Medical Assistant", 1.3d, 56000 );
        dataTable.Rows.Add( "Carl Jones", "Web Designer", 4d, 66000 );
        dataTable.Rows.Add( "Anna Karlweiss", "Account Executive", 7.8d, 51000 );
        dataTable.Rows.Add( "Julia Robertson", "Marketing Coordinator", 17.6d, 65000 );
        var dataTableImportOptions = new DataTableImportOptions() { DestinationTopLeftAddress = "B5", ColumnNames = new string[] { "Name", "Experience", "Position" }, IsColumnNamesShown = true };
        worksheet.ImportData( dataTable, dataTableImportOptions );

        // AutoFit all columns, starting at row 4 up to row 10 (rowId starts at 0), with column sizes from 0 to 255.
        worksheet.Columns.AutoFit( 0, 255, 4, 10 );

        // Center data in column C.
        worksheet.Columns[ "C" ].Style.Alignment.Horizontal = HorizontalAlignment.Center;
        // Bold DataTable's ColumnNames.
        worksheet.Rows[ 4 ].Style.Font.Bold = true;
        // Create a Table with the DataTable.
        var table = worksheet.Tables.Add( "B5", "D10" );
        table.AutoFilter.ShowFilterButton = false;

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: ImportDataTables.xslx\n" );
      }
    }

    public static void ImportCSV()
    {
      using( var workbook = Workbook.Create( ImportDataSample.ImportDataSampleOutputDirectory + @"ImportCSV.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Import CSV";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        worksheet.Cells[ "B4" ].Value = "Import a CSV from a string:";
        worksheet.Cells[ "B4" ].Style.Font.Color = Color.Blue;

        // Define a path to a csv document, the import options(which separator to use) and call the ImportData function.
        var stringSCVData = ImportDataSample.ImportDataSampleResourcesDirectory + @"Book1.csv";
        var stringCSVImportOptions = new CSVImportOptions() { DestinationTopLeftAddress = "C5", Separator = "," };
        worksheet.ImportData( stringSCVData, stringCSVImportOptions );


        worksheet.Cells[ "B11" ].Value = "Import a CSV from a stream:";
        worksheet.Cells[ "B11" ].Style.Font.Color = Color.Blue;

        // Define a stream from a csv document, the import options(which separator to use) and call the ImportData function.
        var streamCSVData = new MemoryStream();
        var file = new FileStream( ImportDataSample.ImportDataSampleResourcesDirectory + @"Book1.csv", FileMode.Open, FileAccess.Read );
        var bytes = new byte[ file.Length ];
        file.Read( bytes, 0, ( int )file.Length );
        streamCSVData.Write( bytes, 0, ( int )file.Length );

        var streamCSVImportOptions = new CSVImportOptions() { DestinationTopLeftAddress = "C12", Separator = "," };
        worksheet.ImportData( stringSCVData, streamCSVImportOptions );

        // Center data in columns 2 to 10 (columnId starts at 0).
        worksheet.Columns[ 2, 10 ].Style.Alignment.Horizontal = HorizontalAlignment.Center;
       
        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: ImportCSV.xslx\n" );
      }
    }

    #endregion

    #region Private classes

    private enum Team
    {
      Chicago_Hornets,
      Miami_Ducks,
      NewYork_Bucs,
      LosAngelese_Raiders
    }

    private class Player
    {
      public string Name { get; set; }

      public int Number { get; set; }

      public Team Team { get; set; }
    }

    #endregion
  }
}
