/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Hyperlink Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to work with hyperlinks when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.IO;

namespace Xceed.Workbooks.NET.Examples
{
  public class HyperlinkSample
  {
    #region Private Members

    private const string HyperlinkSampleResourcesDirectory = Program.SampleDirectory + @"Hyperlink\Resources\";
    private const string HyperlinkSampleOutputDirectory = Program.SampleDirectory + @"Hyperlink\Output\";

    #endregion

    #region Constructors

    static HyperlinkSample()
    {
      if( !Directory.Exists( HyperlinkSample.HyperlinkSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( HyperlinkSample.HyperlinkSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void AddHyperlink()
    {
      using ( var workbook = Workbook.Create( HyperlinkSample.HyperlinkSampleOutputDirectory + @"AddHyperlink.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Add Hyperlinks";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Add an hyperlink for a cell reference(z1) in the same worksheet.
        // Position the hyperlink in cell B6 and extend the hyperlink on 2 columns.
        worksheet.Cells[ "B5" ].Value = "Add an hyperlink to a cell reference in same worksheet:";
        worksheet.Hyperlinks.Add( "z1", "B6", 1, 2, "Link to another cell.", "A cell reference." );

        // Add a new worksheet and fill cells.
        workbook.Worksheets.Add();
        workbook.Worksheets[ 1 ].Cells[ "D4" ].Value = "The other worksheet.";

        // Add an hyperlink for a cell reference in another worksheet (sheet2, cell B1).
        // Position the hyperlink in 8th row and 2nd column ans extend the hyperlink on 3 columns.
        worksheet.Cells[ "B8" ].Value = "Add an hyperlink to a cell reference in another worksheet:";
        worksheet.Hyperlinks.Add( "Sheet2!B1", 8, 1, 1, 3, "Link to another worksheet's cell.", "Another worksheet cell reference." );

        // Add an hyperlink to an external document.
        // Position the hyperlink in cell B12 and extend the hyperlink to cell D12.
        worksheet.Cells[ "B11" ].Value = "Add an hyperlink to an external document:";
        worksheet.Hyperlinks.Add( "../../Worksheet/Output/CalculateWorksheetFormulas.xlsx", "B12", "D12", "Link to another document.", "An external document link." );

        // Add an hyperlink to an email address.
        // Position the hyperlink in cell B15 and extend the hyperlink for 2 rows and 2 columns.
        worksheet.Cells[ "B14" ].Value = "Add an hyperlink to an email address:";
        worksheet.Hyperlinks.Add( "sales@xceed.com", "B15", 2, 2, null, "An email link." );

        // Add an hyperlink to a web site.
        // Position the hyperlink in cell B18.
        worksheet.Cells[ "B17" ].Value = "Add an hyperlink to a web site:";
        worksheet.Hyperlinks.Add( "www.xceed.com", "B18", 1, 1, "Xceed", "A web site link." );

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: AddHyperlink.xslx\n" );
      }
    }

    #endregion
  }
}
