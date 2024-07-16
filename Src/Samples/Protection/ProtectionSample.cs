/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Protection Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to work with protection when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.IO;

namespace Xceed.Workbooks.NET.Examples
{
  public class ProtectionSample
  {
    private const string ProtectionSampleResourcesDirectory = Program.SampleDirectory + @"Protection\Resources\";
    private const string ProtectionSampleOutputDirectory = Program.SampleDirectory + @"Protection\Output\";

    static ProtectionSample()
    {
      if( !Directory.Exists( ProtectionSample.ProtectionSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( ProtectionSample.ProtectionSampleOutputDirectory );
      }
    }

    public static void AddWorksheetProtection()
    {
      using( var workbook = Workbook.Create( ProtectionSampleOutputDirectory + "AddWorksheetProtection.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Add Worksheet Protection"; 
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set content.
        worksheet.Cells[ "C6" ].Value = "This worksheet is protected, no action can be done.";
        worksheet.Cells[ "C8" ].Value = "The worksheet can be unprotected through the 'Review-Unprotect sheet' option.";

        // Protect the 1st worksheet.
        worksheet.Protect();

        // Add a 2nd worksheet.
        var worksheet2 = workbook.Worksheets.Add();

        // Set content in 2nd worksheet.
        worksheet2.Cells[ "B5" ].Value = "This worksheet is NOT protected.";

        workbook.Save();
        Console.WriteLine( "\tCreated: AddWorksheetProtection.xlsx\n" );
      }
    }

    public static void AddWorksheetProtectionWithPassword()
    {
      using( var workbook = Workbook.Create( ProtectionSampleOutputDirectory + "AddWorksheetProtectionWithPassword.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Add Worksheet Protection with password";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set content.
        worksheet.Cells[ "C6" ].Value = "This worksheet is protected with a password, no action can be done.";
        worksheet.Cells[ "C8" ].Value = "The worksheet can be unprotected through the 'Review-Unprotect sheet' option by typing 'xceed'.";

        // Protect the 1st worksheet.
        worksheet.Protect( null, "xceed" );

        // Add a 2nd worksheet.
        var worksheet2 = workbook.Worksheets.Add();

        // Set content in 2nd worksheet.
        worksheet2.Cells[ "B5" ].Value = "This worksheet is NOT protected.";

        workbook.Save();
        Console.WriteLine( "\tCreated: AddWorksheetProtectionWithPassword.xlsx\n" );
      }
    }

    public static void AddWorksheetProtectionAndAllowActions()
    {
      using( var workbook = Workbook.Create( ProtectionSampleOutputDirectory + "AddWorksheetProtectionAndAllowActions.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Add Worksheet Protection and allow actions";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set content.
        worksheet.Cells[ "C6" ].Value = "This worksheet is protected with a password. Only formatting cell and inserting rows/columns is allowed.";
        worksheet.Cells[ "C8" ].Value = "The worksheet can be unprotected through the 'Review-Unprotect sheet' option by typing 'xceed'.";

        // Set the worksheet protection : only formatting cells and inserting rows/columns will be allowed.
        var protection = new WorksheetProtection() { AllowFormatCells = true, AllowInsertRows = true, AllowInsertColumns = true };
        // Protect the 1st worksheet with a password.
        worksheet.Protect( protection, "xceed" );

        // Add a 2nd worksheet.
        var worksheet2 = workbook.Worksheets.Add();

        // Set content in 2nd worksheet.
        worksheet2.Cells[ "B5" ].Value = "This worksheet is NOT protected.";

        workbook.Save();
        Console.WriteLine( "\tCreated: AddWorksheetProtectionAndAllowActions.xlsx\n" );
      }
    }

    public static void RemoveWorksheetProtection()
    {
      using( var workbook = Workbook.Load( ProtectionSampleResourcesDirectory + "RemoveWorksheetProtection.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Remove the protection on the 1st worksheet with a password.
        worksheet.Unprotect( "xceed" );

        workbook.SaveAs( ProtectionSampleOutputDirectory + "RemoveWorksheetProtection.xlsx" );
        Console.WriteLine( "\tCreated: RemoveWorksheetProtection.xlsx\n" );
      }
    }

    public static void UnlockSpecificCells()
    {
      using( var workbook = Workbook.Create( ProtectionSampleOutputDirectory + "UnlockSpecificCells.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Unlock Specific Cells";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set content.
        worksheet.Cells[ "B4" ].Value = "Only Light Green cells are unlocked.";
        worksheet.Cells[ "B4" ].Style.Font.Bold = true;

        worksheet.Cells[ "C6" ].Value = "Date";
        worksheet.Cells[ "D6" ].Value = "Employee";
        worksheet.Cells[ "E6" ].Value = "In Time";
        worksheet.Cells[ "F6" ].Value = "Out Time";

        worksheet.Cells[ "C7" ].Value = new DateTime( 2022, 5, 1 );
        worksheet.Cells[ "D7" ].Value = "Micheal Smith";
        worksheet.Cells[ "E7" ].Value = new TimeSpan( 8, 0, 0 );
        worksheet.Cells[ "F7" ].Value = new TimeSpan( 15, 30, 0 );
        worksheet.Cells[ "D8" ].Value = "Stella Corleone";
        worksheet.Cells[ "E8" ].Value = new TimeSpan( 9, 15, 0 );
        worksheet.Cells[ "F8" ].Value = new TimeSpan(  16, 30, 0 );

        worksheet.Cells[ "C10" ].Value = new DateTime( 2022, 5, 2 );
        worksheet.Cells[ "D10" ].Value = "Carl Debrusk";        
        worksheet.Cells[ "E10" ].Value = new TimeSpan( 8, 15, 0 );
        worksheet.Cells[ "F10" ].Value = new TimeSpan( 13, 45, 0 );
        worksheet.Cells[ "D11" ].Value = "Stella Corleone";
        worksheet.Cells[ "E11" ].Value = new TimeSpan( 8, 45, 0 );
        worksheet.Cells[ "F11" ].Value = new TimeSpan( 12, 30, 0 );
        worksheet.Cells[ "D12" ].Value = "Michael Smith";

        // AutoFit columns from rowId 6 until the end for width between 0 and 255. RowId starts at 0.
        worksheet.Columns.AutoFit( 0, 255, 6 );
        // Format cells in column "E" and "F" to display times.
        worksheet.Columns[ "E", "F" ].Style.CustomFormat = "hh:mm";
        // Horitonaly align content in columns "C" through "F".
        worksheet.Columns[ "C", "F" ].Style.Alignment.Horizontal = HorizontalAlignment.Center;
        // Put the data in a formatted table.
        var table = worksheet.Tables.Add( "C6", "F12", TableStyle.TableStyleLight14 );
        table.AutoFilter.ShowFilterButton = false;

        // All cells have their "locked" property set to true by default and will be activated when the worksheet will be protected.
        // Unlock cells from "E7" to "F12" and set a LightGreen background.
        worksheet.Cells[ "E7", "F12" ].Style.Protection.Locked = false;
        worksheet.Cells[ "E7", "F12" ].Style.Fill.BackgroundColor = System.Drawing.Color.LightGreen;

        // Protect the worksheet.
        worksheet.Protect();

        workbook.Save();
        Console.WriteLine( "\tCreated: UnlockSpecificCells.xlsx\n" );
      }
    }

    public static void LockSpecificCells()
    {
      using( var workbook = Workbook.Create( ProtectionSampleOutputDirectory + "LockSpecificCells.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Lock Specific Cells";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set content.
        worksheet.Cells[ "B4" ].Value = "Only Light Pink cells are locked.";
        worksheet.Cells[ "B4" ].Style.Font.Bold = true;

        worksheet.Cells[ "C6" ].Value = "Date";
        worksheet.Cells[ "D6" ].Value = "Employee";
        worksheet.Cells[ "E6" ].Value = "In Time";
        worksheet.Cells[ "F6" ].Value = "Out Time";

        worksheet.Cells[ "C7" ].Value = new DateTime( 2022, 5, 1 );
        worksheet.Cells[ "D7" ].Value = "Micheal Smith";
        worksheet.Cells[ "E7" ].Value = new TimeSpan( 8, 0, 0 );
        worksheet.Cells[ "F7" ].Value = new TimeSpan( 15, 30, 0 );
        worksheet.Cells[ "D8" ].Value = "Stella Corleone";
        worksheet.Cells[ "E8" ].Value = new TimeSpan( 9, 15, 0 );
        worksheet.Cells[ "F8" ].Value = new TimeSpan( 16, 30, 0 );

        worksheet.Cells[ "C10" ].Value = new DateTime( 2022, 5, 2 );
        worksheet.Cells[ "D10" ].Value = "Carl Debrusk";
        worksheet.Cells[ "E10" ].Value = new TimeSpan( 8, 15, 0 );
        worksheet.Cells[ "F10" ].Value = new TimeSpan( 13, 45, 0 );
        worksheet.Cells[ "D11" ].Value = "Stella Corleone";
        worksheet.Cells[ "E11" ].Value = new TimeSpan( 8, 45, 0 );
        worksheet.Cells[ "F11" ].Value = new TimeSpan( 12, 30, 0 );
        worksheet.Cells[ "D12" ].Value = "Michael Smith";

        // AutoFit columns from rowId 6 until the end for width between 0 and 255. RowId starts at 0.
        worksheet.Columns.AutoFit( 0, 255, 6 );
        // Format cells in column "E" and "F" to display times.
        worksheet.Columns[ "E", "F" ].Style.CustomFormat = "hh:mm";
        // Horitonaly align content in columns "C" through "F".
        worksheet.Columns[ "C", "F" ].Style.Alignment.Horizontal = HorizontalAlignment.Center;
        // Put the data in a formatted table.
        var table = worksheet.Tables.Add( "C6", "F12", TableStyle.TableStyleLight14 );
        table.AutoFilter.ShowFilterButton = false;

        // All cells have their "locked" property set to true by default and will be activated when the worksheet will be protected.
        // So, we unlock the first 100 worksheet's columns cells, so they can be edited.
        worksheet.Columns[ 0, 100 ].Style.Protection.Locked = false;

        // Lock cells from "C7" to "F8", hide them for the formula bar and set a LightPink background.
        worksheet.Cells[ "C7", "F8" ].Style.Protection.Locked = true;
        worksheet.Cells[ "C7", "F8" ].Style.Protection.HiddenFormula = true;
        worksheet.Cells[ "C7", "F8" ].Style.Fill.BackgroundColor = System.Drawing.Color.LightPink;

        // Protect the worksheet and only allow inserting new rows. Unlocked cells can also be edited.
        var protection = new WorksheetProtection() { AllowInsertRows = true };
        worksheet.Protect( protection );

        workbook.Save();
        Console.WriteLine( "\tCreated: LockSpecificCells.xlsx\n" );
      }
    }
  }
}
