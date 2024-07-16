/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Cell Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to work with cells when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.IO;

namespace Xceed.Workbooks.NET.Examples
{
  public class CellSample
  {
    #region Private Members

    private const string CellSampleResourcesDirectory = Program.SampleDirectory + @"Cell\Resources\";
    private const string CellSampleOutputDirectory = Program.SampleDirectory + @"Cell\Output\";

    private enum EnumValues
    {
      Enum_1,
      Enum_2,
      Enum_3
    }

    #endregion

    #region Constructors

    static CellSample()
    {
      if( !Directory.Exists( CellSample.CellSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( CellSample.CellSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void SetCellValueTypes()
    {
      using( var workbook = Workbook.Create( CellSample.CellSampleOutputDirectory + @"SetCellValueTypes.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Set Cell Value Types";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set a value of type number in cells from 2nd column. Indexing starts at 0 for rows and columns.
        worksheet.Cells[ 3, 0 ].Value = "Numeric types";
        worksheet.Cells[ 4, 0 ].Value = "int:";
        worksheet.Cells[ 4, 1 ].Value = ( int )25;
        worksheet.Cells[ 5, 0 ].Value = "double:";
        worksheet.Cells[ 5, 1 ].Value = ( double )33.54;
        worksheet.Cells[ 6, 0 ].Value = "float:";
        worksheet.Cells[ 6, 1 ].Value = ( float )4.5;
        worksheet.Cells[ 7, 0 ].Value = "decimal:";
        worksheet.Cells[ 7, 1 ].Value = ( decimal )22.586;
        worksheet.Cells[ 8, 0 ].Value = "short:";
        worksheet.Cells[ 8, 1 ].Value = ( short )55;
        worksheet.Cells[ 9, 0 ].Value = "long:";
        worksheet.Cells[ 9, 1 ].Value = ( long )8465;
        worksheet.Cells[ 10, 0 ].Value = "byte:";
        worksheet.Cells[ 10, 1 ].Value = ( byte )255;
        worksheet.Cells[ 11, 0 ].Value = "uint:";
        worksheet.Cells[ 11, 1 ].Value = ( uint )152;
        worksheet.Cells[ 12, 0 ].Value = "ulong:";
        worksheet.Cells[ 12, 1 ].Value = ( ulong )101234;
        worksheet.Cells[ 13, 0 ].Value = "ushort:";
        worksheet.Cells[ 13, 1 ].Value = ( ushort )128;
        worksheet.Cells[ 14, 0 ].Value = "sbyte:";
        worksheet.Cells[ 14, 1 ].Value = ( sbyte )-128;

        // Create a table with the numeric typed cells.
        CellSample.CreateFormattedTable( worksheet, 3, 0, 14, 1 );

        // Set a value of type Date in cells from 5th column. Indexing starts at 0 for rows and columns.
        worksheet.Cells[ 3, 3 ].Value = "Date/Time types";
        worksheet.Cells[ 4, 3 ].Value = "DateTime:";
        worksheet.Cells[ 4, 4 ].Value = DateTime.Now;
        worksheet.Cells[ 5, 3 ].Value = "TimeSpan:";
        worksheet.Cells[ 5, 4 ].Value = new TimeSpan( 2, 1, 25, 32 );

        // Create a table with the dateTime typed cells.
        CellSample.CreateFormattedTable( worksheet, 3, 3, 5, 4 );

        // Set a value of type Text in cells from 8th column. Indexing starts at 0 for rows and columns.
        worksheet.Cells[ 3, 6 ].Value = "Text types";
        worksheet.Cells[ 4, 6 ].Value = "string:";
        worksheet.Cells[ 4, 7 ].Value = "This is a string";
        worksheet.Cells[ 5, 6 ].Value = "enum:";
        worksheet.Cells[ 5, 7 ].Value = EnumValues.Enum_1;
        worksheet.Cells[ 6, 6 ].Value = "char:";
        worksheet.Cells[ 6, 7 ].Value = 'c';
        worksheet.Cells[ 7, 6 ].Value = "guid:";
        worksheet.Cells[ 7, 7 ].Value = Guid.NewGuid();

        // Create a table with the text typed cells.
        CellSample.CreateFormattedTable( worksheet, 3, 6, 7, 7 );

        // Set a value of type boolean in cells from 11th column. Indexing starts at 0 for rows and columns.
        worksheet.Cells[ 3, 9 ].Value = "Boolean types";
        worksheet.Cells[ 4, 9 ].Value = "bool:";
        worksheet.Cells[ 4, 10 ].Value = true;
        worksheet.Cells[ 5, 9 ].Value = "bool:";
        worksheet.Cells[ 5, 10 ].Value = false;

        // Center align all the cells of the 4th row.
        worksheet.Rows[ 3 ].Style.Alignment.Horizontal = HorizontalAlignment.Center;

        // Create a table with the boolean typed cells.
        CellSample.CreateFormattedTable( worksheet, 3, 9, 5, 10 );

        // AutoFit all the columns with content starting at the 4th row, and make sure the column's widths are between 0 and 255.
        worksheet.Columns.AutoFit( 0, 255, 3 );        

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: SetCellValueTypes.xlsx\n" );
      }
    }

    public static void SetFormulas()
    {
      using( var workbook = Workbook.Load( CellSample.CellSampleResourcesDirectory + @"CellData.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Set Formulas";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        worksheet.Cells[ "I4" ].Value = "Stats displayed with Tables:";
        worksheet.Cells[ "I4" ].Style.Font = new Font() { Bold = true };

        // Set stats and formulas for Jane.
        // Formulas will be calculated when opening saved bookmark with MS Excel or by calling worksheet.CalculateFormulas().
        worksheet.Cells[ "I6" ].Value = "Stats for Jane:";
        worksheet.Cells[ "I7" ].Value = "Total(2011):";
        worksheet.Cells[ "J7" ].Formula = "=SUM( G6, G12, G18, G24 )";
        worksheet.Cells[ "I8" ].Value = "Total(2012):";
        worksheet.Cells[ "J8" ].Formula = "=SUM( G19, G22 )";
        worksheet.Cells[ "I9" ].Value = "Total(2013):";
        worksheet.Cells[ "J9" ].Formula = "=SUM( G8, G17, G20, G23, G26 )";
        worksheet.Cells[ "I10" ].Value = "Total:";
        worksheet.Cells[ "J10" ].Formula = "=SUM( J7:J9 )";
        worksheet.Cells[ "I11" ].Value = "Avg:";
        worksheet.Cells[ "J11" ].Formula = "=AVERAGE( J7:J9 )";
        worksheet.Cells[ "I13" ].Value = "Bonus:";
        worksheet.Cells[ "J13" ].Formula = "=AVERAGE( J11 * 15% )";

        // Create a table with the Jane stat cells.
        CellSample.CreateFormattedTable( worksheet, "I6", "J13" );

        // Set stats and formulas for Ashish.
        // Formulas will be calculated when opening saved bookmark with MS Excel or by calling worksheet.CalculateFormulas().
        worksheet.Cells[ "L6" ].Value = "Stats for Ashish:";
        worksheet.Cells[ "L7" ].Value = "Total(2011):";
        worksheet.Cells[ "M7" ].Formula = "=SUM( G15 )";
        worksheet.Cells[ "L8" ].Value = "Total(2012):";
        worksheet.Cells[ "M8" ].Formula = "=SUM( G7, G10, G13, G16 )";
        worksheet.Cells[ "L9" ].Value = "Total(2013):";
        worksheet.Cells[ "M9" ].Formula = "=SUM( 0 )";
        worksheet.Cells[ "L10" ].Value = "Total:";
        worksheet.Cells[ "M10" ].Formula = "=SUM( M7:M9 )";
        worksheet.Cells[ "L11" ].Value = "Avg:";
        worksheet.Cells[ "M11" ].Formula = "=AVERAGE( M7:M9 )";
        worksheet.Cells[ "L13" ].Value = "Bonus:";
        worksheet.Cells[ "M13" ].Formula = "=AVERAGE( M11 * 15% )";

        // Create a table with the Ashish stat cells.
        CellSample.CreateFormattedTable( worksheet, "L6", "M13" );

        // Set stats and formulas for John.
        // Formulas will be calculated when opening saved bookmark with MS Excel or by calling worksheet.CalculateFormulas().
        worksheet.Cells[ "O6" ].Value = "Stats for John:";
        worksheet.Cells[ "O7" ].Value = "Total(2011):";
        worksheet.Cells[ "P7" ].Formula = "=SUM( G9, G21 )";
        worksheet.Cells[ "O8" ].Value = "Total(2012):";
        worksheet.Cells[ "P8" ].Formula = "=SUM( G25 )";
        worksheet.Cells[ "O9" ].Value = "Total(2013):";
        worksheet.Cells[ "P9" ].Formula = "=SUM( G11, G14 )";
        worksheet.Cells[ "O10" ].Value = "Total:";
        worksheet.Cells[ "P10" ].Formula = "=SUM( P7:P9 )";
        worksheet.Cells[ "O11" ].Value = "Avg:";
        worksheet.Cells[ "P11" ].Formula = "=AVERAGE( P7:P9 )";
        worksheet.Cells[ "O13" ].Value = "Bonus:";
        worksheet.Cells[ "P13" ].Formula = "=AVERAGE( P11 * 15% )";

        // Create a table with the John stat cells.
        CellSample.CreateFormattedTable( worksheet, "O6", "P13" );

        // Set Formatting for columns.
        worksheet.Columns[ "J" ].Style.CustomFormat = "$#,##0";
        worksheet.Columns[ "M" ].Style.CustomFormat = "$#,##0";
        worksheet.Columns[ "P" ].Style.CustomFormat = "$#,##0";

        // This call will calculate all the formulas of the current worksheet and set the corresponding Cell.Value.
        // It should only be called if the result of calculation is needed before saving the workbook.
        // All formulas will be automatically calculated when opening a workbook with MS Excel.
        // Here we use it because we need the cell.Value's width to autoFit the columns.
        worksheet.CalculateFormulas();

        // AutoFit the columns "I" to "P", from 6th and going down, and make sure the column's widths are between 0 and 255.
        worksheet.Columns[ "I","P" ].AutoFit( 0, 255, 5);

        // Save workbook to disk.
        workbook.SaveAs( CellSample.CellSampleOutputDirectory + @"SetFormulas.xlsx" );
        Console.WriteLine( "\tCreated: SetFormulas.xlsx\n" );
      }
    }

    public static void MergeCells()
    {
      using( var workbook = Workbook.Create( CellSample.CellSampleOutputDirectory + @"MergeCells.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells["B1"].Value = "Merge and unmerge cells";
        worksheet.Cells["B1"].Style.Font = new Font() { Bold = true, Size = 15.5d };

        //Merging only keep starting element informations.
        worksheet.Cells["B3"].Value = "Centered merge using cells addresses";
        worksheet.Cells["B3"].Style.Font = new Font() { Underline = true, Size = 12.5d };
        worksheet.Cells["B4"].Value = "Some centered text.";
        worksheet.MergedCells.Add("B4", "C5");

        worksheet.Cells[6, 1].Value = "Centered merge using IDs";
        worksheet.Cells[6, 1].Style.Font = new Font() { Underline = true, Size = 12.5d };
        worksheet.Cells[7, 1].Value = "Some other text.";
        worksheet.MergedCells.Add(7, 1, 8, 2);

        //Centered using a cell range
        worksheet.Cells[10, 1].Value = "Centered merge using a cell range.";
        worksheet.Cells[10, 1].Style.Font = new Font() { Underline = true, Size = 12.5d };
        worksheet.Cells[11, 1].Value = "Another text.";
        worksheet.Cells[11, 1, 12, 2].MergeCells();

        //Uncentered across using cell range
        worksheet.Cells[14, 1].Value = "Uncentered across merge using a cell range.";
        worksheet.Cells[14, 1].Style.Font = new Font() { Underline = true, Size = 12.5d };
        worksheet.Cells[15, 1].Value = "Here is a text in across merge.";
        worksheet.Cells[16, 1].Value = "Here is another text in across merge.";
        worksheet.Cells[15, 1, 16, 4].MergeCells( false, true ) ;

        //Remove using cell range
        var cellRange = worksheet.Cells[17, 1, 18, 2];
        cellRange.MergeCells();
        cellRange.UnmergeCells();

        //Remove using the MergedCellCollection
        cellRange.MergeCells();
        worksheet.MergedCells.RemoveAt(5);

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: MergeCells.xlsx\n" );
      }
    }

    public static void CellWithMultipleFont()
    {
      using( var workbook = Workbook.Create( CellSample.CellSampleOutputDirectory + @"CellWithMultipleFont.xlsx" ) )
      {
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        FormattedText formattedText = new FormattedText( "Text with multiple font", new Font() { Bold = true, Size = 15.5d } );
        worksheet.Cells[ "B1" ].Value = formattedText ;

        // Create multiple text with different font
        FormattedText formattedText1 = new FormattedText( "This is the first part of the cell ", new Font() { Italic = true, Size = 14, Color = System.Drawing.Color.Aquamarine } );
        FormattedText formattedText2 = new FormattedText( "Here is a normal string " );
        FormattedText formattedText3 = new FormattedText( "Here is another independant text",
          new Font() { Bold = true, Underline = true, UnderlineType = UnderlineType.Double, Color = System.Drawing.Color.Coral} );

        //Put all the text in a list
        FormattedTextList formattedTextsList = new FormattedTextList { formattedText1, formattedText2, formattedText3 };
        
        //Assign value to the desired cell
        worksheet.Cells[ "B3" ].Value = formattedTextsList;

        //Save the workbook
        workbook.SaveAs( CellSample.CellSampleOutputDirectory + @"CellWithMultipleFont.xlsx" );
        Console.WriteLine( "\tCreated: CellWithMultipleFont.xlsx\n" );
      }
    }
    
    public static void FormatPartOfText()
    {
      using( var workbook = Workbook.Create( CellSample.CellSampleOutputDirectory + @"FormatPartOfText.xlsx" ) )
      {
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        FormattedText formattedText = new FormattedText( "Reformat a part of an existing text", new Font() { Bold = true, Size = 15.5d } );
        worksheet.Cells[ "B1" ].Value = formattedText ;

        // Create multiple text with different font
        FormattedText formattedText1 = new FormattedText( "This is the first part of the cell ", new Font() { Italic = true, Size = 14, Color = System.Drawing.Color.Aquamarine } );
        FormattedText formattedText2 = new FormattedText( "Here is a normal string " );
        FormattedText formattedText3 = new FormattedText( "Here is another independant text",
          new Font() { Bold = true, Underline = true, UnderlineType = UnderlineType.Double, Color = System.Drawing.Color.Coral} );

        //Put all the text in a list
        FormattedTextList formattedTextsList = new FormattedTextList { formattedText1, formattedText2, formattedText3 };
        
        //Assign value to the desired cell
        worksheet.Cells[ "B3" ].Value = formattedTextsList;

        worksheet.Cells[ "B3" ].FormatText( new Font(){Color = System.Drawing.Color.Blue}, 5, 10 );

        //Save the workbook
        workbook.SaveAs( CellSample.CellSampleOutputDirectory + @"FormatPartOfText.xlsx" );
        Console.WriteLine( "\tCreated: FormatPartOfText.xlsx\n" );
      }
    }

    public static void ReplaceContent()
    {
      using( var workbook = Workbook.Load( CellSample.CellSampleResourcesDirectory + @"CellData.xlsx" ) )
      {
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Replace cell's content";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d }; 

        // Replace all occurences of string "Jane" with "Michael".
        worksheet.ReplaceContent( "Jane", "Michael" );

        // Add formatted texts in cell "I6".
        var formattedTexts = new FormattedTextList()
        {
          new FormattedText( "* Name ", new Font() { Size = 13 } ),
          new FormattedText( "Jane", new Font() { Color = System.Drawing.Color.Red, Bold = true, Size = 13 } ),
          new FormattedText( " has been replaced with ", new Font() { Size = 13 } ),
          new FormattedText( "Michael", new Font() { Color = System.Drawing.Color.Green, Bold = true, Size = 13 } )
        };
        worksheet.Cells[ "I6" ].Value = formattedTexts;

        //Save the workbook
        workbook.SaveAs( CellSample.CellSampleOutputDirectory + @"ReplaceContent.xlsx" );
        Console.WriteLine( "\tCreated: ReplaceContent.xlsx\n" );
      }
    }

    public static void DeleteCellRange()
    {
      using( var workbook = Workbook.Create( CellSample.CellSampleOutputDirectory + @"DeleteCellRange.xlsx" ) )
      {
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Delete cell Range";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill Cells for initial display
        worksheet.Cells[ "C5" ].Value = "Before Delete cellRange";
        worksheet.Cells[ "C5" ].Style.Font = new Font() { Bold = true };
        worksheet.Cells[ "C6" ].Value = 1;
        worksheet.Cells[ "D6" ].Value = 2;
        worksheet.Cells[ "E6" ].Value = 3;
        worksheet.Cells[ "C7" ].Value = 4;
        worksheet.Cells[ "D7" ].Value = 5;
        worksheet.Cells[ "E7" ].Value = 6;
        worksheet.Cells[ "C8" ].Value = 7;
        worksheet.Cells[ "D8" ].Value = 8;
        worksheet.Cells[ "E8" ].Value = 9;
        worksheet.Cells[ "C9" ].Value = 10;
        worksheet.Cells[ "D9" ].Value = 11;
        worksheet.Cells[ "E9" ].Value = 12;

        // Fill Cells for resulting display
        worksheet.Cells[ "H5" ].Value = "After Delete cellRange (I6 to I7)";
        worksheet.Cells[ "H5" ].Style.Font = new Font() { Bold = true };
        worksheet.Cells[ "H6" ].Value = 1;
        worksheet.Cells[ "I6" ].Value = 2;
        worksheet.Cells[ "J6" ].Value = 3;
        worksheet.Cells[ "H7" ].Value = 4;
        worksheet.Cells[ "I7" ].Value = 5;
        worksheet.Cells[ "J7" ].Value = 6;
        worksheet.Cells[ "H8" ].Value = 7;
        worksheet.Cells[ "I8" ].Value = 8;
        worksheet.Cells[ "J8" ].Value = 9;
        worksheet.Cells[ "H9" ].Value = 10;
        worksheet.Cells[ "I9" ].Value = 11;
        worksheet.Cells[ "J9" ].Value = 12;

        // Delete CellRange from I6 to I7 and shift the following cells up.
        worksheet.DeleteRange( "I6", "I7", DeleteRangeShiftType.ShiftCellsUp );

        //Save the workbook
        workbook.Save();
        Console.WriteLine( "\tCreated: DeleteCellRange.xlsx\n" );
      }
    }

    public static void InsertCellRange()
    {
      using( var workbook = Workbook.Create( CellSample.CellSampleOutputDirectory + @"InsertCellRange.xlsx" ) )
      {
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Insert cell Range";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Fill Cells for initial display
        worksheet.Cells[ "C5" ].Value = "Before Insert cellRange";
        worksheet.Cells[ "C5" ].Style.Font = new Font() { Bold = true };
        worksheet.Cells[ "C6" ].Value = 1;
        worksheet.Cells[ "D6" ].Value = 2;
        worksheet.Cells[ "E6" ].Value = 3;
        worksheet.Cells[ "C7" ].Value = 4;
        worksheet.Cells[ "D7" ].Value = 5;
        worksheet.Cells[ "E7" ].Value = 6;
        worksheet.Cells[ "C8" ].Value = 7;
        worksheet.Cells[ "D8" ].Value = 8;
        worksheet.Cells[ "E8" ].Value = 9;
        worksheet.Cells[ "C9" ].Value = 10;
        worksheet.Cells[ "D9" ].Value = 11;
        worksheet.Cells[ "E9" ].Value = 12;

        // Fill Cells for resulting display
        worksheet.Cells[ "H5" ].Value = "After Insert cellRange (I7 to I8)";
        worksheet.Cells[ "H5" ].Style.Font = new Font() { Bold = true };
        worksheet.Cells[ "H6" ].Value = 1;
        worksheet.Cells[ "I6" ].Value = 2;
        worksheet.Cells[ "J6" ].Value = 3;
        worksheet.Cells[ "H7" ].Value = 4;
        worksheet.Cells[ "I7" ].Value = 5;
        worksheet.Cells[ "J7" ].Value = 6;
        worksheet.Cells[ "H8" ].Value = 7;
        worksheet.Cells[ "I8" ].Value = 8;
        worksheet.Cells[ "J8" ].Value = 9;
        worksheet.Cells[ "H9" ].Value = 10;
        worksheet.Cells[ "I9" ].Value = 11;
        worksheet.Cells[ "J9" ].Value = 12;

        // Insert CellRange from I7 to I8 and shift the following cells down.
        worksheet.InsertRange( "I7", "I8", InsertRangeShiftType.ShiftCellsDown );

        //Save the workbook
        workbook.Save();
        Console.WriteLine( "\tCreated: InsertCellRange.xlsx\n" );
      }
    }

    #endregion

    #region Private Methods

    private static void CreateFormattedTable( Worksheet worksheet, int startRowId, int startColumnId, int endRowId, int endColumnId )
    {
      var table = worksheet.Tables.Add( startRowId, startColumnId, endRowId, endColumnId, TableStyle.TableStyleMedium9 );
      table.ShowFirstColumnFormatting = true;
      table.AutoFilter.ShowFilterButton = false;
      table.Columns[ 1 ].Name = "Values";
    }

    private static void CreateFormattedTable( Worksheet worksheet, string startAddress, string endAddress )
    {
      var table = worksheet.Tables.Add( startAddress, endAddress, TableStyle.TableStyleMedium9 );
      table.ShowFirstColumnFormatting = true;
      table.AutoFilter.ShowFilterButton = false;
      table.Columns[ 1 ].Name = "Values";
    }

    #endregion
  }
}
