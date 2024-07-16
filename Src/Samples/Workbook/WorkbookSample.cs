/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Workbook Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to create or load an Excel workbook when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Xceed.Workbooks.NET.Examples
{
  public class WorkbookSample
  {
    #region Private Members

    private const string WorkbookSampleResourcesDirectory = Program.SampleDirectory + @"Workbook\Resources\";
    private const string WorkbookSampleOutputDirectory = Program.SampleDirectory + @"Workbook\Output\";

    #endregion

    #region Constructors

    static WorkbookSample()
    {
      if( !Directory.Exists( WorkbookSample.WorkbookSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( WorkbookSample.WorkbookSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void CreateWorkbook()
    {
      using( var workbook = Workbook.Create( WorkbookSample.WorkbookSampleOutputDirectory + @"CreateWorkbook.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Create a Workbook";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Generate stats.
        var rnd = new Random();
        var northDivisionTeams = new Dictionary<string, Data>()
        {
          { "Montreal Canadiens", WorkbookSample.GetStats(rnd) },
          { "Toronto Maple Leafs", WorkbookSample.GetStats(rnd) },
          { "Edmonton Oilers", WorkbookSample.GetStats(rnd) },
          { "Calgary Flames", WorkbookSample.GetStats(rnd) },
          { "Winnipeg Jets", WorkbookSample.GetStats(rnd) },
          { "Ottawa Senators", WorkbookSample.GetStats(rnd) },
          { "Vancouver Canucks", WorkbookSample.GetStats(rnd) }
        };

        // Fill cells values and row styles/size.
        worksheet.Rows[ 4 ].Cells[ 2 ].Value = "2021 NHL Standings";
        worksheet.Rows[ 4 ].Style.Font.Bold = true;
        worksheet.Rows[ 4 ].Height = 50;

        worksheet.Rows[ 5 ].Cells[ 0 ].Value = "North Division";
        worksheet.Rows[ 6 ].Cells[ 1 ].Value = "Pts";
        worksheet.Rows[ 6 ].Cells[ 2 ].Value = "Wins";
        worksheet.Rows[ 6 ].Cells[ 3 ].Value = "%";
        worksheet.Rows[ 6 ].Cells[ 4 ].Value = "In Playoffs";
        worksheet.Rows[ 6 ].Cells[ 5 ].Value = "Last Win";
        worksheet.Rows[ 6 ].Style.Alignment.Horizontal = HorizontalAlignment.Center;

        var teams = northDivisionTeams.OrderByDescending( entry => entry.Value.Pts );
        for( int i = 0; i < teams.Count(); ++i )
        {
          var team = teams.ElementAt( i );
          worksheet.Rows[ 7 + i ].Cells[ 0 ].Value = team.Key;
          worksheet.Rows[ 7 + i ].Cells[ 1 ].Value = team.Value.Pts;
          worksheet.Rows[ 7 + i ].Cells[ 2 ].Value = team.Value.Wins;
          worksheet.Rows[ 7 + i ].Cells[ 3 ].Value = team.Value.Percent;
          worksheet.Rows[ 7 + i ].Cells[ 4 ].Value = ( i <= 3 );
          worksheet.Rows[ 7 + i ].Cells[ 5 ].Value = team.Value.LastWin;
        }

        // Set the style display format for all cells in column "D".
        worksheet.Columns[ "D" ].Style.CustomFormat = "0.000";

        // AutoFit all columns with content, from row 6 to 13, with a minimum width of 0 and a maximum width of 255.
        worksheet.Columns.AutoFit( 0, 255, 6, 13 );

        // Set Outline and inside borders for CellRange A7 to F14.
        worksheet.Cells[ "A7", "F14" ].Style.Borders.SetOutline( LineStyle.Medium, System.Drawing.Color.Blue );
        worksheet.Cells[ "A7", "F14" ].Style.Borders.SetInside( LineStyle.Medium, System.Drawing.Color.Blue );

        // Set Fill for CellRange A7 to F7
        worksheet.Cells[ "A7", "F7" ].Style.Fill.BackgroundColor = System.Drawing.Color.Orange;

        // Save the created workbook.
        workbook.Save();
        Console.WriteLine( "\tCreated: CreateWorkbook.xlsx\n" );
      }
    }

    public static void LoadWorkbookWithFilename()
    {
      using( var workbook = Workbook.Load( WorkbookSample.WorkbookSampleResourcesDirectory + @"AutoValue.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Load Workbook with filname";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Insert cell values into this workbook.
        worksheet.Cells[ "C27" ].Value = "Manager:";
        worksheet.Cells[ "D27" ].Value = "Mike Thompson";
        // Set 27th row font.
        worksheet.Rows[ 26 ].Style.Font.Bold = true;

        // Save workbook to disk.
        workbook.SaveAs( WorkbookSample.WorkbookSampleOutputDirectory + @"LoadWorkbookWithFilename.xlsx" );
        Console.WriteLine( "\tCreated: LoadWorkbookWithFilename.xlsx\n" );
      }
    }

    public static void LoadWorkbookWithStream()
    {
      using( var fs = new FileStream( WorkbookSample.WorkbookSampleResourcesDirectory + @"AutoValue.xlsx", FileMode.Open, FileAccess.Read, FileShare.Read ) )
      {
        using( var workbook = Workbook.Load( fs ) )
        {
          // Get the first worksheet. A workbook contains at least 1 worksheet.
          var worksheet = workbook.Worksheets[ 0 ];

          // Add a title.
          worksheet.Cells[ "B1" ].Value = "Load Workbook with stream";
          worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

          // Insert cell values into this workbook.
          worksheet.Cells[ "C27" ].Value = "Manager:";
          worksheet.Cells[ "D27" ].Value = "Mike Thompson";
          // Set 27th row font.
          worksheet.Rows[ 26 ].Style.Font.Bold = true;

          // Save workbook to disk.
          workbook.SaveAs( WorkbookSample.WorkbookSampleOutputDirectory + @"LoadWorkbookWithStream.xlsx" );
          Console.WriteLine( "\tCreated: LoadWorkbookWithStream.xlsx\n" );
        }
      }
    }

    public static void LoadWorkbookWithStringUrl()
    {
      using( var workbook = Workbook.Load( "https://trumpexcel.com/wp-content/uploads/2015/09/Excel-To-Do-List-Template-Print.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Fill cell values.
        worksheet.Cells[ "C4" ].Value = "Paint wall";
        worksheet.Cells[ "D4" ].Value = "In Red";

        // Save workbook to disk.
        workbook.SaveAs( WorkbookSample.WorkbookSampleOutputDirectory + @"LoadWorkbookWithUrl.xlsx" );
        Console.WriteLine( "\tCreated: LoadWorkbookWithUrl.xlsx\n" );
      }
    }

    public static void CalculateWorkbookFormulas()
    {
      using( var workbook = Workbook.Create( WorkbookSample.WorkbookSampleOutputDirectory + @"CalculateWorkbookFormulas.xlsx" ) )
      {
        // Add a second worksheet in workbook.
        workbook.Worksheets.Add();

        var worksheetA = workbook.Worksheets[ 0 ];
        var worksheetB = workbook.Worksheets[ 1 ];

        // Add a title.
        worksheetA.Cells[ "B1" ].Value = "Calculate Workbook Formulas";
        worksheetA.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill cells values in WorksheetA.
        worksheetA.Cells[ "A5" ].Value = "Employees";
        worksheetA.Cells[ "B5" ].Value = "Salary";
        worksheetA.Cells[ "A6" ].Value = "Mike Jones";
        worksheetA.Cells[ "B6" ].Value = 52000;
        worksheetA.Cells[ "A7" ].Value = "Cathy Smith";
        worksheetA.Cells[ "B7" ].Value = 46000;
        worksheetA.Cells[ "A8" ].Value = "Kevin Malcolm";
        worksheetA.Cells[ "B8" ].Value = 77000;
        worksheetA.Cells[ "A9" ].Value = "Jenny McIntyre";
        worksheetA.Cells[ "B9" ].Value = 61000;
        worksheetA.Cells[ "A10" ].Value = "AVERAGE:";

        // AutoFit first column in WorksheetA.
        worksheetA.Columns["A"].AutoFit();

        // Set Font for 5th and 10th row.
        worksheetA.Rows[ 4 ].Style.Font.Bold = true;
        worksheetA.Rows[ 9 ].Style.Font.Bold = true;

        // Set second column's format in WorksheetA and width.
        worksheetA.Columns[ "B" ].Style.CustomFormat = "$#,###";
        worksheetA.Columns[ "B" ].Width = 12d;

        // Set average employees salary formula in WorksheetA.
        worksheetA.Cells[ "B10" ].Formula = "=AVERAGE(B6:B9)";

        // Fill cells values in WorksheetB.
        worksheetB.Cells[ "A1" ].Value = "Item number";
        worksheetB.Cells[ "B1" ].Value = "Screws required";
        worksheetB.Cells[ "A2" ].Value = "G017";
        worksheetB.Cells[ "B2" ].Value = 22;
        worksheetB.Cells[ "A3" ].Value = "K147";
        worksheetB.Cells[ "B3" ].Value = 32;
        worksheetB.Cells[ "A4" ].Value = "A689";
        worksheetB.Cells[ "B4" ].Value = 12;
        worksheetB.Cells[ "A5" ].Value = "B127";
        worksheetB.Cells[ "B5" ].Value = 16;
        worksheetB.Cells[ "A6" ].Value = "TOTAL:";

        // AutoFit all columns in WorksheetB.
        worksheetB.Columns.AutoFit();

        // Set Font for 1st and 6th row.
        worksheetB.Rows[ 0 ].Style.Font.Bold = true;
        worksheetB.Rows[ 5 ].Style.Font.Bold = true;

        // Set total screws required formula in WorksheetB.
        worksheetB.Cells[ "B6" ].Formula = "=SUM(B2:B5)";

        // Cells with formula do not set their Value property until opened with MS Excel or CalculateFormulas() is called.
        var worksheetA_averageValue = worksheetA.Cells[ "B10" ].Value;
        var worksheetB_sumValue = worksheetB.Cells[ "B6" ].Value;
        Debug.Assert( worksheetA_averageValue == null );
        Debug.Assert( worksheetB_sumValue == null );

        // Calculate formulas for all worksheets.
        workbook.CalculateFormulas();

        // Cells with formula now have their Value property calculated .
        worksheetA_averageValue = worksheetA.Cells[ "B10" ].Value;
        worksheetB_sumValue = worksheetB.Cells[ "B6" ].Value;
        Debug.Assert( worksheetA_averageValue != null );
        Debug.Assert( worksheetB_sumValue != null );

        // Display calculation results in other cells.
        worksheetA.Cells[ "C15" ].Value = "Result of formula calculation for Sheet1 is:  " + worksheetA_averageValue;
        worksheetA.Cells[ "C16" ].Value = "Result of formula calculation for Sheet2 is:  " + worksheetB_sumValue;

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: CalculateWorkbookFormulas.xlsx\n" );
      }
    }

    #endregion

    #region Private Methods

    static Data GetStats( Random rnd )
    {
      var wins = rnd.Next( 0, 57 );
      var pts = ( wins * 2 ) + rnd.Next( 0, 56 - wins ) / 2;
      var percent = Convert.ToDouble( wins ) / 56d;
      var lastWin = new DateTime( 2021, rnd.Next( 3, 5 ), rnd.Next( 1, 31 ) );

      return new Data() 
      { Pts = pts, 
        Wins = wins, 
        Percent = percent, 
        LastWin = lastWin
      };
    }

    #endregion

    #region Private Classes

    private struct Data
    {
      public int Pts;
      public int Wins;
      public double Percent;
      public DateTime LastWin;
    };

    #endregion
  }
}
