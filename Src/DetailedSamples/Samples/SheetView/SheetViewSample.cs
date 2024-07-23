/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – SheetView Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to work with sheetViews when using the API 
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
  public class SheetViewSample
  {
    #region Private Members

    private const string SheetViewSampleResourcesDirectory = Program.SampleDirectory + @"SheetView\Resources\";
    private const string SheetViewSampleOutputDirectory = Program.SampleDirectory + @"SheetView\Output\";

    #endregion

    #region Constructors

    static SheetViewSample()
    {
      if( !Directory.Exists( SheetViewSample.SheetViewSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( SheetViewSample.SheetViewSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void SetActiveCell()
    {
      using( var workbook = Workbook.Load( SheetViewSample.SheetViewSampleResourcesDirectory + @"Sheet.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Set Active and TopLeft Cells";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        worksheet.Cells[ "B4" ].Value = "Using a TopLeftCell at B1 and an ActiveCell at E20:";
        worksheet.Cells[ "B4" ].Style.Font = new Font() { Bold = true };

        // Set the active cell and top left cell of the 1st worksheet.
        worksheet.SheetView.ActiveCellAddress = "E20";
        worksheet.SheetView.TopLeftCellAddress = "B1";

        // Save workbook to disk.
        workbook.SaveAs( SheetViewSample.SheetViewSampleOutputDirectory + @"SetActiveCell.xlsx" );
        Console.WriteLine( "\tCreated: SetActiveCell.xlsx\n" );
      }
    }

    public static void SetZoomAndViewType()
    {
      using( var workbook = Workbook.Load( SheetViewSample.SheetViewSampleResourcesDirectory + @"Sheet.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Set Zoom and ViewType";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Modify the Normal View scale zoom to 82%.
        worksheet.SheetView.ZoomScale = 82;

        // Modify the PageLayout View scale zoom to 166% and set the SheetView type to PageLayout.
        worksheet.SheetView.ZoomScalePageLayout = 166;
        worksheet.SheetView.ViewType = WorksheetViewType.PageLayout;

        // Save workbook to disk.
        workbook.SaveAs( SheetViewSample.SheetViewSampleOutputDirectory + @"SetZoomAndViewType.xlsx" );
        Console.WriteLine( "\tCreated: SetZoomAndViewType.xlsx\n" );
      }
    }

    public static void FreezeRowsColumns()
    {
      using( var workbook = Workbook.Create( SheetViewSample.SheetViewSampleResourcesDirectory + @"FreezeRowsColumns.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Freeze Rows and Columns";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill cells content.
        SheetViewSample.FillCellContent( worksheet );

        // Freeze horizontally after first 5 rows and add a Fill Background on them.
        worksheet.SheetView.FrozenRows = 5;
        worksheet.Rows[ 0, 4 ].Style.Fill.BackgroundColor = System.Drawing.Color.Turquoise;

        // Freeze vertically after first Column and add a Fill Background on it.
        worksheet.SheetView.FrozenColumns = 1;
        worksheet.Columns[ 0 ].Style.Fill.BackgroundColor = System.Drawing.Color.Tan;

        // Save workbook to disk.
        workbook.SaveAs( SheetViewSample.SheetViewSampleOutputDirectory + @"FreezeRowsColumns.xlsx" );
        Console.WriteLine( "\tCreated: FreezeRowsColumns.xlsx\n" );
      }
    }

    public static void SplitRowsColumns()
    {
      using( var workbook = Workbook.Create( SheetViewSample.SheetViewSampleResourcesDirectory + @"SplitRowsColumns.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Split Rows and Columns";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill cells content.
        SheetViewSample.FillCellContent( worksheet );

        // Split horizontally after first 5 rows and add a Fill Background on them.
        worksheet.SheetView.SplitRows = 5;
        worksheet.Rows[ 0, 4 ].Style.Fill.BackgroundColor = System.Drawing.Color.LightCyan;

        // Split vertically after 1st column and add a Fill Background on first column.
        worksheet.SheetView.SplitColumns = 1;
        worksheet.Columns[ 0 ].Style.Fill.BackgroundColor = System.Drawing.Color.MediumSpringGreen;

        // Save workbook to disk.
        workbook.SaveAs( SheetViewSample.SheetViewSampleOutputDirectory + @"SplitRowsColumns.xlsx" );
        Console.WriteLine( "\tCreated: SplitRowsColumns.xlsx\n" );
      }
    }

    #endregion

    #region Private Methods

    private static void FillCellContent( Worksheet worksheet )
    {
      var random = new Random();

      Debug.Assert( worksheet != null, "Worksheet whouldn't be null.");

      for( var columnId = 1; columnId <= 50; ++columnId )
      {
        // First Rows of data cells.
        worksheet.Cells[ 4, columnId ].Value = new DateTime( 2022, 6, 1 ).AddDays( columnId - 1 );
      }
      for( var rowId = 5; rowId < 53; ++rowId )
      {
        // First Columns of data cells.
        worksheet.Cells[ rowId, 0 ].Value = new DateTime( 2022, 6, 1 ).AddMinutes( ( rowId - 5 ) * 30 );
      }
      for( var columnId = 1; columnId <= 50; ++columnId )
      {
        for( var rowId = 5; rowId < 53; ++rowId )
        {
          // Inner cells.
          worksheet.Cells[ rowId, columnId ].Value = random.Next( 0, 2 ) == 1 ? "YES" : "NO";
        }
      }

      // Format DateTimes.
      worksheet.Rows[ 4 ].Style.CustomFormat = "yyyy/MM/dd";
      worksheet.Columns[ 0 ].Style.CustomFormat = "HH:mm";

      // Center cell content rowId, columnId from (4, 0) to (52, 50)
      worksheet.Cells[ 4, 0, 52, 50 ].Style.Alignment.Horizontal = HorizontalAlignment.Center;

      // AutoFit Columns, with a width from 0 to 255, starting at rowId 4.
      worksheet.Columns.AutoFit( 0, 255, 4 );
    }

    #endregion
  }
}
