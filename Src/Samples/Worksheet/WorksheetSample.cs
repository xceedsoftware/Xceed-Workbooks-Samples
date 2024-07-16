/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Worksheet Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to work with worksheets when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;

namespace Xceed.Workbooks.NET.Examples
{
  public class WorksheetSample
  {
    #region Private Members

    private const string WorksheetSampleResourcesDirectory = Program.SampleDirectory + @"Worksheet\Resources\";
    private const string WorksheetSampleOutputDirectory = Program.SampleDirectory + @"Worksheet\Output\";

    #endregion

    #region Constructors

    static WorksheetSample()
    {
      if( !Directory.Exists( WorksheetSample.WorksheetSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( WorksheetSample.WorksheetSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void AddWorksheets()
    {
      using( var workbook = Workbook.Create( WorksheetSample.WorksheetSampleOutputDirectory + @"AddWorksheets.xlsx" ) )
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

        // We now have 3 Worksheets in the Workbook.
        Debug.Assert( workbook.Worksheets.Count == 3 );
        Debug.Assert( workbook.Worksheets.Contains( "Third Sheet" ) );

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: AddWorksheets.xlsx\n" );
      }
    }

    public static void RemoveWorksheets()
    {
      using( var workbook = Workbook.Create( WorksheetSample.WorksheetSampleOutputDirectory + @"RemoveWorksheets.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Remove Worksheets";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill cells in 1st worksheet.
        worksheet.Cells[ "D5" ].Value = "This is the first Worksheet.";

        // Add a worksheet with default "Sheet2" name.
        workbook.Worksheets.Add();

        // Fill cells in 2nd worksheet.
        workbook.Worksheets[ 1 ].Cells[ "D5" ].Value = "This is the second Worksheet.";

        // Add a worksheet with name "Third Sheet".
        workbook.Worksheets.Add( "Third Sheet" );

        // Fill cells in 3rd worksheet.
        workbook.Worksheets[ "Third Sheet" ].Cells[ "D5" ].Value = "This is the third Worksheet.";

        // Add a worksheet with default "Sheet4" name.
        workbook.Worksheets.Add();

        // Fill cells in 4th worksheet.
        workbook.Worksheets[ 3 ].Cells[ "D5" ].Value = "This is the fourth Worksheet.";

        // We now have 4 Worksheets in the Workbook.
        Debug.Assert( workbook.Worksheets.Count == 4 );

        // Remove all worksheets from worksheet id = 1 to worksheet id = 2.
        workbook.Worksheets.Remove( workbook.Worksheets[ 1, 2 ] );

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: RemoveWorksheets.xlsx\n" );
      }
    }

    public static void HideWorksheets()
    {
      using( var workbook = Workbook.Create( WorksheetSample.WorksheetSampleOutputDirectory + @"HideWorksheets.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Hide Worksheets";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill cells in 1st worksheet.
        worksheet.Cells[ "D5" ].Value = "This is the First Worksheet.";

        // Add 3 worksheets and fill them.
        workbook.Worksheets.Add();
        workbook.Worksheets[ 1 ].Cells[ "D5" ].Value = "This is the Second Worksheet.";
        workbook.Worksheets.Add();
        workbook.Worksheets[ 2 ].Cells[ "D5" ].Value = "This is the Third Worksheet.";
        workbook.Worksheets.Add();
        workbook.Worksheets[ 3 ].Cells[ "D5" ].Value = "This is the Fourth Worksheet.";
        workbook.Worksheets.Add();
        workbook.Worksheets[ 4 ].Cells[ "D5" ].Value = "This is the Fifth Worksheet.";

        // We now have 4 Worksheets in the Workbook.
        Debug.Assert( workbook.Worksheets.Count == 5 );

        // Hide the 2nd worksheet. It can be unhided through MS Excel UI.
        workbook.Worksheets[ 1 ].Visibility = WorksheetVisibility.Hidden;

        // Hide the 3rd and 4th worksheets. They can NOT be unhided through MS Excel UI.
        workbook.Worksheets[ 2, 3 ].Visibility = WorksheetVisibility.Collapsed;

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: HideWorksheets.xlsx\n" );
      }
    }

    public static void CellAccess()
    {
      using( var workbook = Workbook.Create( WorksheetSample.WorksheetSampleOutputDirectory + @"CellAccess.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Cell Access";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set a value in cell located at 4th row and 2nd column. Indexing starts at 0 for rows and columns.
        // Accessing with indexes could be faster with many cell access.
        worksheet.Cells[ 3, 1 ].Value = "Cell value at row 3 and column 2";

        // Set a value in cell located at address "D8". Indexing starts at "A" for columns and "1" for rows.
        // Accessing with addresses could be slower with many cell access.
        worksheet.Cells[ "D8" ].Value = "Cell value at address D8";

        // AutoFit all columns with values.
        worksheet.Columns.AutoFit();

        // Making sure only 3 cells in the worksheet exists (the modified cells).
        Debug.Assert( worksheet.Cells.Count == 3 );

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: CellAccess.xlsx\n" );
      }
    }

    public static void ColumnAccess()
    {
      using( var workbook = Workbook.Create( WorksheetSample.WorksheetSampleOutputDirectory + @"ColumnAccess.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Column Access";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set the width of the 3rd column. Indexing with numbers starts at 0 for columns.
        worksheet.Columns[ 2 ].Width = 60d;

        // Set the format of the 5th column. Indexing with letters starts at "A" for columns.
        worksheet.Columns[ "E" ].Style.CustomFormat = "$0.000";

        // Making sure only 3 columns exist in the worksheet (the modified columns or columns with cell values).
        Debug.Assert( worksheet.Columns.Count == 3 );

        // Set values in cells.
        worksheet.Columns[ 2 ].Cells[ 5 ].Value = "This column has a width of 60.";
        worksheet.Columns[ "E" ].Cells[ 8 ].Value = "A formatted column";
        worksheet.Columns[ "E" ].Cells[ 10 ].Value = 58.364215;

        // AutoFit column "E".
        worksheet.Columns[ "E" ].AutoFit();

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: ColumnAccess.xlsx\n" );
      }
    }

    public static void RowAccess()
    {
      using( var workbook = Workbook.Create( WorksheetSample.WorksheetSampleOutputDirectory + @"RowAccess.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Row Access";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set the Height of the 5th row. Indexing with numbers starts at 0 for rows.
        worksheet.Rows[ 4 ].Height = 30d;

        // Making sure only 2 rows exist in the worksheet (the modified rows or rows with cell values).
        Debug.Assert( worksheet.Rows.Count == 2 );

        // Set values in cells.
        worksheet.Rows[ 4 ].Cells[ 5 ].Value = "This row has a height of 30.";

        // AutoFit for 6th column.
        worksheet.Columns[ 5 ].AutoFit();

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: RowAccess.xlsx\n" );
      }
    }

    public static void CustomizeWorksheets()
    {
      // Load a workbook.
      using( var workbook = Workbook.Load( WorksheetSample.WorksheetSampleResourcesDirectory + @"ThreeWorksheets.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheetA = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheetA.Cells[ "B1" ].Value = "Customize Worksheets";
        worksheetA.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set the name and tabColor for first worksheet.
        worksheetA.Name = "2019";
        worksheetA.TabColor = Color.Red;

        // Gets 2nd Worksheet and set its name and tabColor.
        var worksheetB = workbook.Worksheets[ "Sheet2" ];
        worksheetB.Name = "2020";
        worksheetB.TabColor = Color.FromArgb( 0, 255, 0 );

        // Set the active cell and top left cell in the 2nd worksheet.
        worksheetB.SheetView.ActiveCellAddress = "E20";
        worksheetB.SheetView.TopLeftCellAddress = "D18";

        // Gets 3rd Worksheet and set its name and tabColor.
        var worksheetC = workbook.Worksheets[ 2 ];
        worksheetC.Name = "2021";
        worksheetC.TabColor = Color.Blue;

        // Save workbook to disk.
        workbook.SaveAs( WorksheetSample.WorksheetSampleOutputDirectory + @"CustomizeWorksheets.xlsx" );
        Console.WriteLine( "\tCreated: CustomizeWorksheets.xlsx\n" );
      }
    }

    public static void CalculateWorksheetFormulas()
    {
      using( var workbook = Workbook.Create( WorksheetSample.WorksheetSampleOutputDirectory + @"CalculateWorksheetFormulas.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Calculate Worksheet Formulas";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill cells values in WorksheetA.
        worksheet.Cells[ "A5" ].Value = "Employees";
        worksheet.Cells[ "B5" ].Value = "Salary";
        worksheet.Cells[ "A6" ].Value = "Mike Jones";
        worksheet.Cells[ "B6" ].Value = 52000;
        worksheet.Cells[ "A7" ].Value = "Cathy Smith";
        worksheet.Cells[ "B7" ].Value = 46000;
        worksheet.Cells[ "A8" ].Value = "Kevin Malcolm";
        worksheet.Cells[ "B8" ].Value = 77000;
        worksheet.Cells[ "A9" ].Value = "Jenny McIntyre";
        worksheet.Cells[ "B9" ].Value = 61000;
        worksheet.Cells[ "A10" ].Value = "AVERAGE:";

        // AutoFit first column in WorksheetA.
        worksheet.Columns[ "A" ].AutoFit();

        // Set second column's format in WorksheetA.
        worksheet.Columns[ "B" ].Style.CustomFormat = "$#,###";

        // Set rows styles for 5th and 10th rows.
        worksheet.Rows[ 4 ].Style.Font.Bold = true;
        worksheet.Rows[ 9 ].Style.Font.Bold = true;

        // Set average employees salary formula in Worksheet.
        worksheet.Cells[ "B10" ].Formula = "=AVERAGE(B6:B9)";

        // Cells with formula do not set their Value property until opened with MS Excel or CalculateFormulas() is called.
        var worksheet_averageValue = worksheet.Cells[ "B10" ].Value;
        Debug.Assert( worksheet_averageValue == null );

        // Calculate formulas for this worksheet.
        worksheet.CalculateFormulas();

        // Cells with formula now have a calculated value.
        worksheet_averageValue = worksheet.Cells[ "B10" ].Value;
        Debug.Assert( worksheet_averageValue != null );

        // Display calculation results in other cells.
        worksheet.Cells[ "C15" ].Value = "Result of formula calculation for Sheet1 in B10 is:  " + worksheet_averageValue;

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: CalculateWorksheetFormulas.xlsx\n" );
      }
    }

    public static void CopyWorksheet()
    {
      using( var workbook = Workbook.Create( WorksheetSample.WorksheetSampleOutputDirectory + @"CopyWorksheet.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Copy Worksheet";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set the Height of the 5th row. Indexing with numbers starts at 0 for rows.
        worksheet.Rows[ 4 ].Height = 30d;

        // Set values in cells.
        worksheet.Rows[ 4 ].Cells[ 5 ].Value = "This row has a height of 30.";

        // AutoFit for 6th column.
        worksheet.Columns[ 5 ].AutoFit();

        // Copy the first worksheet and named it "The new name".
        // The worksheet to copy can be an id, a name or an instance.
        var worksheetCopy = workbook.Worksheets.Copy( 0, "The new name" );

        // Input new values in copied worksheet.
        worksheetCopy.Cells[ "A10" ].Value = "This is the copied worksheet.";

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: WorksheetCopy.xslx\n" );
      }
    }

    public static void MoveWorksheets()
    {
      using( var workbook = Workbook.Create( WorksheetSample.WorksheetSampleOutputDirectory + @"MoveWorksheets.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet_A = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet_A.Cells[ "B1" ].Value = "Move Worksheets";
        worksheet_A.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill cells in 1st worksheet.
        worksheet_A.Cells[ "D5" ].Value = "This is the first Worksheet.";

        // Add a worksheet with default "SheetX" name.
        var worksheet_B = workbook.Worksheets.Add();

        // Fill cells in 2nd worksheet.
        worksheet_B.Cells[ "D5" ].Value = "This is the second Worksheet.";

        // Add a worksheet with name "Third Sheet".
        var worksheet_C = workbook.Worksheets.Add( "Third Sheet" );

        // Fill cells in 3rd worksheet.
        worksheet_C.Cells[ "D5" ].Value = "This is the third Worksheet.";

        // Move the third worksheet at index 0 (to be first).
        // The Move function takes worksheet ids, names or instances.
        workbook.Worksheets.Move( worksheet_C.Name, 0 );

        // Assign the active worksheet to be the 2nd one.
        workbook.WorkbookViews[ 0 ].ActiveTab = 1;

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: MoveWorksheets.xslx\n" );
      }
    }

    public static void InsertDeleteRows()
    {
      using( var workbook = Workbook.Create( WorksheetSample.WorksheetSampleOutputDirectory + @"InsertDeleteRows.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Insert and delete Rows";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill cells for first table.
        worksheet.Cells[ "D4" ].Value = "Insert Rows:";
        worksheet.Cells[ "D4" ].Style.Font.Bold = true;
        worksheet.Cells[ "D5" ].Value = "2022 Math class - Students Test Results:";
        worksheet.Cells[ "D5" ].Style.Font.Bold = true;

        worksheet.Cells[ "D6" ].Value = "Name";
        worksheet.Cells[ "E6" ].Value = "Midterm(40%)";
        worksheet.Cells[ "F6" ].Value = "Final(60%)";
        worksheet.Cells[ "D7" ].Value = "Mike Jones";
        worksheet.Cells[ "E7" ].Value = 84;
        worksheet.Cells[ "F7" ].Value = 78;
        worksheet.Cells[ "D8" ].Value = "Kelly Smith";
        worksheet.Cells[ "E8" ].Value = 85;
        worksheet.Cells[ "F8" ].Value = 82;
        worksheet.Cells[ "D9" ].Value = "Cindy Newman";
        worksheet.Cells[ "E9" ].Value = 71;
        worksheet.Cells[ "F9" ].Value = 81;
        worksheet.Cells[ "D10" ].Value = "Michael Sawyer";
        worksheet.Cells[ "E10" ].Value = 61;
        worksheet.Cells[ "F10" ].Value = 66;

        // Create a 4-rows table from previous cells, along with the header row.
        var testTable = worksheet.Tables.Add( "D6", "F10", TableStyle.TableStyleDark1 );
        testTable.AutoFilter.ShowFilterButton = false;

        // Insert 2 rows in the middle of the testTable (at rowId 8). RowId starts at 0.
        worksheet.InsertRows( 8, 2 );

        // testTable has2 more rows, extending from cells D6 to F12.
        Debug.Assert( testTable.CellRange.StartingElement.Address == "D6" );
        Debug.Assert( testTable.CellRange.EndingElement.Address == "F12" );


        // Fill cells for 2nd table.
        worksheet.Cells[ "D14" ].Value = "Delete Rows:";
        worksheet.Cells[ "D14" ].Style.Font.Bold = true;
        worksheet.Cells[ "D15" ].Value = "Matt's owned cars (initially 5 blue rows)";
        worksheet.Cells[ "D15" ].Style.Font.Bold = true;

        worksheet.Cells[ "D16" ].Value = "Type";
        worksheet.Cells[ "E16" ].Value = "Bought in";
        worksheet.Cells[ "F16" ].Value = "Sold in";
        worksheet.Cells[ "D17" ].Value = "Chevrolet Cavalier";
        worksheet.Cells[ "E17" ].Value = "June 1999";
        worksheet.Cells[ "F17" ].Value = "September 2004";
        worksheet.Cells[ "D18" ].Value = "Honda Civic";
        worksheet.Cells[ "E18" ].Value = "October 2004";
        worksheet.Cells[ "F18" ].Value = "April 2010";
        worksheet.Cells[ "D19" ].Value = "Toyota Echo";
        worksheet.Cells[ "E19" ].Value = "April 2010";
        worksheet.Cells[ "F19" ].Value = "June 2010";
        worksheet.Cells[ "D20" ].Value = "Dodge Caravan";
        worksheet.Cells[ "E20" ].Value = "July 2010";
        worksheet.Cells[ "F20" ].Value = "March 2018";
        worksheet.Cells[ "D21" ].Value = "Audi A4";
        worksheet.Cells[ "E21" ].Value = "March 2018";
        worksheet.Cells[ "F21" ].Value = "May 2021";

        // Create a 5-rows table from previous cells, along with the header row.
        var carTable = worksheet.Tables.Add( "D16", "F21", TableStyle.TableStyleDark2 );
        carTable.AutoFilter.ShowFilterButton = false;

        // Delete 1 row in the middle of the carTable (at rowId 19). RowId starts at 0.
        worksheet.DeleteRows( 19 );

        // carTable has 1 less row, extending from cells D16 to F20.
        Debug.Assert( carTable.CellRange.StartingElement.Address == "D16" );
        Debug.Assert( carTable.CellRange.EndingElement.Address == "F20" );

        // AutoFits all Columns.
        worksheet.Columns.AutoFit();

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: InsertDeleteRows.xslx\n" );
      }
    }

    public static void InsertDeleteColumns()
    {
      using( var workbook = Workbook.Create( WorksheetSample.WorksheetSampleOutputDirectory + @"InsertDeleteColumns.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Insert and delete Columns";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill cells for first table.
        worksheet.Cells[ "C4" ].Value = "Insert Columns:";
        worksheet.Cells[ "C4" ].Style.Font.Bold = true;
        worksheet.Cells[ "C5" ].Value = "2021 English class - Students Test Results:";
        worksheet.Cells[ "C5" ].Style.Font.Bold = true;

        worksheet.Cells[ "C6" ].Value = "Name";
        worksheet.Cells[ "D6" ].Value = "Midterm(40%)";
        worksheet.Cells[ "E6" ].Value = "Final(40%)";
        worksheet.Cells[ "C7" ].Value = "Mike Jones";
        worksheet.Cells[ "D7" ].Value = 84;
        worksheet.Cells[ "E7" ].Value = 78;
        worksheet.Cells[ "C8" ].Value = "Kelly Smith";
        worksheet.Cells[ "D8" ].Value = 85;
        worksheet.Cells[ "E8" ].Value = 82;
        worksheet.Cells[ "C9" ].Value = "Cindy Newman";
        worksheet.Cells[ "D9" ].Value = 71;
        worksheet.Cells[ "E9" ].Value = 81;
        worksheet.Cells[ "C10" ].Value = "Michael Sawyer";
        worksheet.Cells[ "D10" ].Value = 61;
        worksheet.Cells[ "E10" ].Value = 66;

        // Create a 4-rows table from previous cells, along with the header row.
        var testTable = worksheet.Tables.Add( "C6", "E10", TableStyle.TableStyleDark11 );
        testTable.AutoFilter.ShowFilterButton = false;

        // Center content for Columns "D" and "E".
        worksheet.Columns[ "D", "E" ].Style.Alignment.Horizontal = HorizontalAlignment.Center;

        // Insert 2 columns in the middle of the testTable (at columnId "E").
        worksheet.InsertColumns( "E", 2 );

        // testTable has 2 more columns, extending from cells C6 to G10.
        Debug.Assert( testTable.CellRange.StartingElement.Address == "C6" );
        Debug.Assert( testTable.CellRange.EndingElement.Address == "G10" );


        // Fill cells for 2nd table.
        worksheet.Cells[ "I4" ].Value = "Delete Columns:";
        worksheet.Cells[ "I4" ].Style.Font.Bold = true;
        worksheet.Cells[ "I5" ].Value = "Matt's owned cars (initially 3 columns)";
        worksheet.Cells[ "I5" ].Style.Font.Bold = true;

        worksheet.Cells[ "I6" ].Value = "Type";
        worksheet.Cells[ "J6" ].Value = "Bought in";
        worksheet.Cells[ "K6" ].Value = "Sold in";
        worksheet.Cells[ "I7" ].Value = "Chevrolet Cavalier";
        worksheet.Cells[ "J7" ].Value = "June 1999";
        worksheet.Cells[ "K7" ].Value = "September 2004";
        worksheet.Cells[ "I8" ].Value = "Honda Civic";
        worksheet.Cells[ "J8" ].Value = "October 2004";
        worksheet.Cells[ "K8" ].Value = "April 2010";
        worksheet.Cells[ "I9" ].Value = "Toyota Echo";
        worksheet.Cells[ "J9" ].Value = "April 2010";
        worksheet.Cells[ "K9" ].Value = "June 2010";
        worksheet.Cells[ "I10" ].Value = "Dodge Caravan";
        worksheet.Cells[ "J10" ].Value = "July 2010";
        worksheet.Cells[ "K10" ].Value = "March 2018";
        worksheet.Cells[ "I11" ].Value = "Audi A4";
        worksheet.Cells[ "J11" ].Value = "March 2018";
        worksheet.Cells[ "K11" ].Value = "May 2021";

        // Create a 5-rows table from previous cells, along with the header row.
        var carTable = worksheet.Tables.Add( "I6", "K11", TableStyle.TableStyleDark3 );
        carTable.AutoFilter.ShowFilterButton = false;

        // Delete 1 column in the middle of the carTable (at column 9). RowId starts at 0.
        worksheet.DeleteColumns( 9 );

        // carTable has 1 less column, extending from cells I6 to J11.
        Debug.Assert( carTable.CellRange.StartingElement.Address == "I6" );
        Debug.Assert( carTable.CellRange.EndingElement.Address == "J11" );

        // AutoFits all Columns.
        worksheet.Columns.AutoFit();

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: InsertDeleteColumns.xslx\n" );
      }
    }

    #endregion
  }
}
