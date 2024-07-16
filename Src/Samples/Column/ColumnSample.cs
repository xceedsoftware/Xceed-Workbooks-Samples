/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Column Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to work with columns when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Xceed.Workbooks.NET.Examples
{
  public class ColumnSample
  {
    #region Private Members

    private const string ColumnSampleResourcesDirectory = Program.SampleDirectory + @"Column\Resources\";
    private const string ColumnSampleOutputDirectory = Program.SampleDirectory + @"Column\Output\";

    #endregion

    #region Constructors

    static ColumnSample()
    {
      if( !Directory.Exists( ColumnSample.ColumnSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ColumnSample.ColumnSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void ColumnCellAccess()
    {
      using( var workbook = Workbook.Create( ColumnSample.ColumnSampleOutputDirectory + @"ColumnCellAccess.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Get 4th column. Indexing starts at 0 for columns.
        var columnD = worksheet.Columns[ 3 ];
        // Get column "F". Indexing starts at "A" for columns.
        var columnF = worksheet.Columns[ "F" ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Cell Access";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set a value in cells located at 4th column. Indexing starts at 0 for columns's cells.
        columnD.Cells[ 1 ].Value = "Cell value for 2nd cell of ColumnId 3";
        columnD.Cells[ 3 ].Value = "Cell value for 4th cell of ColumnId 3";
        columnF.Cells[ 8 ].Value = "Cell value for 9th cell of Column F";

        // Set AutoFit for columns with values.
        worksheet.Columns[ 1 ].AutoFit();
        worksheet.Columns[ 3 ].AutoFit();
        worksheet.Columns[ "F" ].AutoFit();

        // Making sure only 2 cells in the 4th column exists and 1 in column "F" (the modified cells).
        Debug.Assert( columnD.Cells.Count == 2 );
        Debug.Assert( columnF.Cells.Count == 1 );

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: ColumnCellAccess.xlsx\n" );
      }
    }

    public static void CustomizeColumns()
    {
      using( var workbook = Workbook.Create( ColumnSample.ColumnSampleOutputDirectory + @"CustomizeColumns.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Customize Columns";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill Cells and width of columns.  
        worksheet.Cells[ "C5" ].Value = "This column has a width of 45";
        worksheet.Columns[ 2 ].Width = 45d;

        worksheet.Cells[ "D6" ].Value = "This column has an autoFit.";
        worksheet.Columns[ "D" ].AutoFit();

        worksheet.Columns[ 5 ].Cells[ 8 ].Value = "This column has a width of 30";
        worksheet.Columns[ 5 ].Width = 30d;

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: CustomizeColumns.xlsx\n" );
      }
    }

    public static void HideUnhideColumns()
    {
      using( var workbook = Workbook.Create( ColumnSample.ColumnSampleOutputDirectory + @"HideUnhideColumns.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Hide/Unhide Columns";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill Cells and style some cells and columns.
        worksheet.Rows[ 4 ].Cells[ 2 ].Value = "Column D through F are hidden, while Column I through K are visible.";
        worksheet.Rows[ 4 ].Cells[ 2 ].Style.Font.Bold = true;

        // Indexes starts at 0, but at 1 in MS Excel.
        for( int i = 3; i < 6; ++i )
        {
          worksheet.Columns[ i ].Cells[ 10 ].Value = "Hidden";
          worksheet.Columns[ i ].Style.Fill.BackgroundColor = System.Drawing.Color.LightPink;
        }
        for( int i = 8; i < 11; ++i )
        {
          worksheet.Columns[ i ].Cells[ 10 ].Value = "Visible";
          worksheet.Columns[ i ].Style.Fill.BackgroundColor = System.Drawing.Color.LightGreen;
        }

        // Hide Columns 3-5 and 8-10. Indexes starts at 0.
        worksheet.Columns[ 3, 5 ].IsHidden = true;
        worksheet.Columns[ 8, 10 ].IsHidden = true;

        // Unhide Columns 8-10. Indexes starts at 0.
        worksheet.Columns[ 8, 10 ].IsHidden = false;

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: HideUnhideColumns.xlsx\n" );
      }
    }

    public static void ClearColumnContents()
    {
      using( var workbook = Workbook.Create( ColumnSample.ColumnSampleOutputDirectory + @"ClearColumnContents.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Clear Column Contents";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        worksheet.Cells[ 4, 6 ].Value = "Column I's content has been cleared, but it's styling remains.";

        // Fill Cells.
        worksheet.Cells[ 8, 6 ].Value = "First Name";
        worksheet.Cells[ 8, 7 ].Value = "Last Name";
        worksheet.Cells[ 8, 8 ].Value = "Age";
        worksheet.Cells[ 8, 9 ].Value = "Is Working";        

        worksheet.Cells[ 9, 6 ].Value = "Tom";
        worksheet.Cells[ 9, 7 ].Value = "Jones";
        worksheet.Cells[ 9, 8 ].Value = 29;
        worksheet.Cells[ 9, 9 ].Value = true;

        worksheet.Cells[ 10, 6 ].Value = "Stella";
        worksheet.Cells[ 10, 7 ].Value = "Smith";
        worksheet.Cells[ 10, 8 ].Value = 38;
        worksheet.Cells[ 10, 9 ].Value = true;

        worksheet.Cells[ 11, 6 ].Value = "Carl";
        worksheet.Cells[ 11, 7 ].Value = "Oconnor";
        worksheet.Cells[ 11, 8 ].Value = 66;
        worksheet.Cells[ 11, 9 ].Value = false;

        worksheet.Cells[ 12, 6 ].Value = "Brianna";
        worksheet.Cells[ 12, 7 ].Value = "Thompson";
        worksheet.Cells[ 12, 8 ].Value = 47;
        worksheet.Cells[ 12, 9 ].Value = false;

        // Style some cells on row 8, from (8, 6) to (8, 9).
        worksheet.Cells[ 8, 6, 8, 9 ].Style.Fill.BackgroundColor = System.Drawing.Color.LightGreen;
        worksheet.Cells[ 8, 6, 8, 9 ].Style.Font = new Font() { Bold = true, Size = 13 };

        // AutoFit all columns in the worksheet, starting at row 8, with width extending from 0 to 255.
        worksheet.Columns.AutoFit( 0, 255, 8 );

        // Clear only the content of column 8.
        worksheet.Columns[ 8 ].Clear( ClearOptions.Contents );

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: ClearColumnContents.xlsx\n" );
      }
    }

    #endregion
  }
}
