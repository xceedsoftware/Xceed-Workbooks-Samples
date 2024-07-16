/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Table Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to work with Tables when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.IO;

namespace Xceed.Workbooks.NET.Examples
{
  public class TableSample
  {
    #region Private Members

    private const string TableSampleResourcesDirectory = Program.SampleDirectory + @"Table\Resources\"; 
    private const string TableSampleOutputDirectory = Program.SampleDirectory + @"Table\Output\";

    #endregion

    #region Constructors

    static TableSample()
    {
      if( !Directory.Exists( TableSample.TableSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( TableSample.TableSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void AddFormattedTable()
    {
      using( var workbook = Workbook.Create( TableSample.TableSampleOutputDirectory + @"AddFormattedTable.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Add Formatted Tables";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill cells values for 1st table.
        worksheet.Rows[ 4 ].Cells[ 2 ].Value = "Players";
        worksheet.Rows[ 5 ].Cells[ 2 ].Value = "Tom Sawyer";
        worksheet.Rows[ 6 ].Cells[ 2 ].Value = "Mika Thompson";
        worksheet.Rows[ 7 ].Cells[ 2 ].Value = "Kevin Smith";
        worksheet.Rows[ 4 ].Cells[ 3 ].Value = "Points";
        worksheet.Rows[ 5 ].Cells[ 3 ].Value = 56;
        worksheet.Rows[ 6 ].Cells[ 3 ].Value = 42;
        worksheet.Rows[ 7 ].Cells[ 3 ].Value = 34;

        // Create the 1st formatted table from cell (4,2) to cell (7,3), using a dark theme and showing the formatted header row.
        worksheet.Tables.Add( 4, 2, 7, 3, TableStyle.TableStyleDark10, true );


        // Fill cells values for 2nd table.
        worksheet.Rows[ 4 ].Cells[ 6 ].Value = "Fruits";
        worksheet.Rows[ 5 ].Cells[ 6 ].Value = "Strawberry";
        worksheet.Rows[ 6 ].Cells[ 6 ].Value = "Orange";
        worksheet.Rows[ 7 ].Cells[ 6 ].Value = "Banana";
        worksheet.Rows[ 8 ].Cells[ 6 ].Value = "Apple";
        worksheet.Rows[ 4 ].Cells[ 7 ].Value = "Quantity";
        worksheet.Rows[ 5 ].Cells[ 7 ].Value = 8;
        worksheet.Rows[ 6 ].Cells[ 7 ].Value = 2;
        worksheet.Rows[ 7 ].Cells[ 7 ].Value = 3;
        worksheet.Rows[ 8 ].Cells[ 7 ].Value = 3;

        // Create the 2nd formatted table from cell (4,6) to cell (8,7), using a medium theme and showing the formatted header row.
        var secondTable = worksheet.Tables.Add( 4, 6, 8, 7, TableStyle.TableStyleMedium20, true );
        // Get this 2nd formatted table and set some properties. Tables and TableColumns can be accessed by id or name.        
        secondTable.ShowFirstColumnFormatting = true;
        secondTable.ShowBandedRows = true;
        worksheet.Tables[ 1 ].AutoFilter.ShowFilterButton = false;
        worksheet.Tables[ "Table2" ].ShowTotalRow = true;
        worksheet.Tables[ "Table2" ].Columns[ "Fruits" ].TotalRowLabel = "TOTAL:";
        worksheet.Tables[ "Table2" ].Columns[ "Quantity" ].TotalRowFunction = TotalRowFunction.Sum;

        // Fill cells values for 3rd table.
        worksheet.Rows[ 4 ].Cells[ 10 ].Value = "Employee";
        worksheet.Rows[ 5 ].Cells[ 10 ].Value = "Nancy Davolio";
        worksheet.Rows[ 6 ].Cells[ 10 ].Value = "Margaret Peacock";
        worksheet.Rows[ 7 ].Cells[ 10 ].Value = "Steven Buchanan";
        worksheet.Rows[ 8 ].Cells[ 10 ].Value = "Laura Callahan";
        worksheet.Rows[ 4 ].Cells[ 11 ].Value = "Years of experience";
        worksheet.Rows[ 5 ].Cells[ 11 ].Value = 12;
        worksheet.Rows[ 6 ].Cells[ 11 ].Value = 2;
        worksheet.Rows[ 7 ].Cells[ 11 ].Value = 9;
        worksheet.Rows[ 8 ].Cells[ 11 ].Value = 7;
        worksheet.Rows[ 4 ].Cells[ 12 ].Value = "Salary";
        worksheet.Rows[ 5 ].Cells[ 12 ].Value = 75000;
        worksheet.Rows[ 6 ].Cells[ 12 ].Value = 41000;
        worksheet.Rows[ 7 ].Cells[ 12 ].Value = 64000;
        worksheet.Rows[ 8 ].Cells[ 12 ].Value = 55000;

        // Adjust alignment and format for 3rd table.
        worksheet.Columns[ 11 ].Style.Alignment.Horizontal = HorizontalAlignment.Center;
        worksheet.Columns[ 12 ].Style.CustomFormat = "$0,000";
        worksheet.Columns[ 12 ].Style.Alignment.Horizontal = HorizontalAlignment.Center;

        // Create the 3rd formatted table from cell (4,10) to cell (8,12), using a light theme and showing the formatted header row.
        var thirdTable = worksheet.Tables.Add( 4, 10, 8, 12, TableStyle.TableStyleLight14, true );
        // Get this 3rd formatted table and set some properties. Tables and TableColumns can be accessed by position or name.        
        thirdTable.ShowFirstColumnFormatting = true;
        thirdTable.ShowLastColumnFormatting = true;
        worksheet.Tables[ 2 ].ShowBandedRows = true;
        worksheet.Tables[ 2 ].ShowBandedColumns = true;
        worksheet.Tables[ 2 ].AutoFilter.ShowFilterButton = false;
        worksheet.Tables[ "Table3" ].ShowTotalRow = true;
        worksheet.Tables[ "Table3" ].Columns[ "Employee" ].TotalRowLabel = "Avg:";
        worksheet.Tables[ "Table3" ].Columns[ "Years of experience" ].TotalRowFunction = TotalRowFunction.Average;
        worksheet.Tables[ "Table3" ].Columns[ "Salary" ].TotalRowFunction = TotalRowFunction.CustomFormula;
        worksheet.Tables[ "Table3" ].Columns[ "Salary" ].TotalRowFormula = "=AVERAGE(M6:M9)";

        // AutoFit all columns with content.
        worksheet.Columns.AutoFit();

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: AddFormattedTable.xlsx\n" );
      }
    }

    public static void RemoveTables()
    {
      using( var workbook = Workbook.Load( TableSample.TableSampleResourcesDirectory + @"Tables.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Modify the title.
        worksheet.Cells[ "B1" ].Value = "Remove Tables";

        // Remove Table at index 1.
        worksheet.Tables.RemoveAt( 1 );
        worksheet.Cells[ "L4" ].Value = "The table has been removed.";

        // Remove Table named "Table4".
        worksheet.Tables.Remove( "Table4" );
        worksheet.Cells[ "L32" ].Value = "The table has been removed.";

        // Set horizontal alignment of column "L" to Left.
        worksheet.Columns[ "L" ].Style.Alignment.Horizontal = HorizontalAlignment.Left;

        // Save workbook to disk.
        workbook.SaveAs( TableSample.TableSampleOutputDirectory + @"RemoveTables.xlsx" );
        Console.WriteLine( "\tCreated: RemoveTables.xlsx\n" );
      }
    }

    #endregion
  }
}
