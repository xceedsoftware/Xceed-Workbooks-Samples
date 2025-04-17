/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to use the different features when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.Drawing;
using System.Reflection;

namespace Xceed.Workbooks.NET.Examples
{
  public class Program
  {
#if NETCORE || NET5
    internal const string SampleDirectory = @"..\..\..\Samples\";
#else
    internal const string SampleDirectory = @"..\..\Samples\";
#endif

    private static void Main( string[] args )
    {
      Xceed.Workbooks.NET.Licenser.LicenseKey = "LICENSE_KEY_PLACEHOLDER";
#if !BLUEPRINT
      XceedDeploymentLicense.SetLicense();
#endif

      var version = Assembly.GetExecutingAssembly().GetName().Version;
      var versionNumber = version.Major + "." + version.Minor;
      Console.WriteLine( "\nRunning Examples of Xceed Workbooks for .NET version " + versionNumber + ".\n" );

      // Workbook
      WorkbookSample.CreateWorkbook();
      WorkbookSample.LoadWorkbookWithFilename();
      WorkbookSample.LoadWorkbookWithStream();
      WorkbookSample.LoadWorkbookWithStringUrl();
      WorkbookSample.CalculateWorkbookFormulas();

      // Worksheet
      WorksheetSample.AddWorksheets();
      WorksheetSample.RemoveWorksheets();
      WorksheetSample.HideWorksheets();
      WorksheetSample.CellAccess();
      WorksheetSample.ColumnAccess();
      WorksheetSample.RowAccess();
      WorksheetSample.CustomizeWorksheets();
      WorksheetSample.CalculateWorksheetFormulas();
      WorksheetSample.CopyWorksheet();
      WorksheetSample.MoveWorksheets();
      WorksheetSample.InsertDeleteRows();
      WorksheetSample.InsertDeleteColumns();

      // Importing Data
      ImportDataSample.ImportArrays();
      ImportDataSample.ImportCollections();
      ImportDataSample.ImportDataTables();
      ImportDataSample.ImportCSV();

      //Hyperlink
      HyperlinkSample.AddHyperlink();

      // SheetView
      SheetViewSample.SetActiveCell();
      SheetViewSample.SetZoomAndViewType();
      SheetViewSample.FreezeRowsColumns();
      SheetViewSample.SplitRowsColumns();

      // Row
      RowSample.RowCellAccess();
      RowSample.CustomizeRows();
      RowSample.HideUnhideRows();

      // Column
      ColumnSample.ColumnCellAccess();
      ColumnSample.CustomizeColumns();
      ColumnSample.HideUnhideColumns();
      ColumnSample.ClearColumnContents();

      // Cell
      CellSample.SetCellValueTypes();
      CellSample.SetFormulas();
      CellSample.MergeCells();
      CellSample.CellWithMultipleFont();
      CellSample.ReplaceContent();
      CellSample.FormatPartOfText();
      CellSample.DeleteCellRange();
      CellSample.InsertCellRange();

      // Tables
      TableSample.AddFormattedTable();
      TableSample.RemoveTables();

      // Miscellaneous
      MiscellaneousSample.LoadDataFromWebToExcel();

      // Picture
      PictureSample.AddPicture();
      PictureSample.OffsetPicture();
      PictureSample.ShrinkPictureWithOffset();

      // Style
      StyleSample.SetAlignments();
      StyleSample.SetFonts();
      StyleSample.SetFills();
      StyleSample.SetBorders();
      StyleSample.SetFormattings();
      StyleSample.SetBuiltinStyles();
      StyleSample.SetStyleOnRanges();
      StyleSample.ModifyTheme();
      StyleSample.ChangeTextDirection();
      StyleSample.ChangeTextOrientation();

      // Protection
      ProtectionSample.AddWorksheetProtection();
      ProtectionSample.AddWorksheetProtectionWithPassword();
      ProtectionSample.AddWorksheetProtectionAndAllowActions();
      ProtectionSample.RemoveWorksheetProtection();
      ProtectionSample.UnlockSpecificCells();
      ProtectionSample.LockSpecificCells();

      // Annotations and Thread Comments
      AnnotationsSample.AddNote();
      AnnotationsSample.AddComment();
      AnnotationsSample.IdentifyNotesOrComments();
      AnnotationsSample.ChangeNoteFormatting();

      Console.WriteLine( "\nDone running Examples of Xceed Workbooks for .NET version " + versionNumber + ".\n" );
      Console.WriteLine( "\nPress any key to exit." );
      Console.ReadKey();
    }
  }
}
