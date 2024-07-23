/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Row Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to work with rows when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.Diagnostics;
using System.IO;

namespace Xceed.Workbooks.NET.Examples
{
  public class RowSample
  {
    #region Private Members

    private const string RowSampleResourcesDirectory = Program.SampleDirectory + @"Row\Resources\";
    private const string RowSampleOutputDirectory = Program.SampleDirectory + @"Row\Output\";

    #endregion

    #region Constructors

    static RowSample()
    {
      if( !Directory.Exists( RowSample.RowSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( RowSample.RowSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void RowCellAccess()
    {
      using( var workbook = Workbook.Create( RowSample.RowSampleOutputDirectory + @"RowCellAccess.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Get 4th row. Indexing starts at 0 for rows.
        var row = worksheet.Rows[ 3 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Cell Access";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set a value in cell located at 4th row. Indexing starts at 0 for row's cells.
        row.Cells[ 1 ].Value = "Cell value for 2nd cell of RowId 3";
        row.Cells[ "D" ].Value = "Cell value for cell in column 'D' of RowId 3";

        // Set AutoFit for columns with values.
        worksheet.Columns[ 1 ].AutoFit();
        worksheet.Columns[ "D" ].AutoFit();

        // Making sure only 2 cells in the 4th row exists (the modified cells).
        Debug.Assert( row.Cells.Count == 2 );

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: RowCellAccess.xlsx\n" );
      }
    }

    public static void CustomizeRows()
    {
      using( var workbook = Workbook.Create( RowSample.RowSampleOutputDirectory + @"CustomizeRows.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];    

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Customize Rows";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill Cells and row Heights.
        worksheet.Rows[ 5 ].Cells[ 3 ].Value = "This row has a height of 30";
        worksheet.Rows[ 5 ].Height = 30d;

        worksheet.Cells[ "B11" ].Value = "This row has a height of 45";
        worksheet.Rows[ 10 ].Height = 45d;

        worksheet.Cells[ "C15" ].Value = "This row has an auto height";
        worksheet.Cells[ "C15" ].Style.Font.Size = 48d;
        worksheet.Rows[ 14 ].AutoFit();

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: CustomizeRows.xlsx\n" );
      }
    }

    public static void HideUnhideRows()
    {
      using( var workbook = Workbook.Create( RowSample.RowSampleOutputDirectory + @"HideUnhideRows.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Hide/Unhide Rows";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill Cells and style some cells and rows.
        worksheet.Rows[ 4 ].Cells[ 2 ].Value = "Row 9 through 11 are hidden, while row 15 through 17 are visible.";
        worksheet.Rows[ 4 ].Cells[ 2 ].Style.Font.Bold = true;

        // Indexes starts at 0, but at 1 in MS Excel.
        for( int i = 8; i < 11; ++i )
        {
          worksheet.Rows[ i ].Cells[ 3 ].Value = "This row is hidden";
          worksheet.Rows[ i ].Style.Fill.BackgroundColor = System.Drawing.Color.LightPink;
        }
        for( int i = 14; i < 17; ++i )
        {
          worksheet.Rows[ i ].Cells[ 3 ].Value = "This row is visible";
          worksheet.Rows[ i ].Style.Fill.BackgroundColor = System.Drawing.Color.LightGreen;
        }

        // Hide Rows 8-10 and 14-16. Indexes starts at 0.
        worksheet.Rows[ 8, 10 ].IsHidden = true;
        worksheet.Rows[ 14, 16 ].IsHidden = true;

        // Unhide Rows 14-16. Indexes starts at 0.
        worksheet.Rows[ 14, 16 ].IsHidden = false;

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: HideUnhideRows.xlsx\n" );
      }
    }

    #endregion
  }
}
