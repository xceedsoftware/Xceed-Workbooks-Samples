/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Style Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to work with styles when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.Drawing;
using System.IO;

namespace Xceed.Workbooks.NET.Examples
{
  public class StyleSample
  {
    #region Private Members

    private const string StyleSampleOutputDirectory = Program.SampleDirectory + @"Style\Output\";

    #endregion

    #region Constructors

    static StyleSample()
    {
      if( !Directory.Exists( StyleSample.StyleSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( StyleSample.StyleSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void SetAlignments()
    {
      using( var workbook = Workbook.Create( StyleSample.StyleSampleOutputDirectory + @"SetAlignments.xlsx" ) )
      {
        // Get first worksheet and change its name.
        var cellWorksheet = workbook.Worksheets[ 0 ];
        cellWorksheet.Name = "Cells";

        // Add a title.
        cellWorksheet.Cells[ "B1" ].Value = "Set Alignments";
        cellWorksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        StyleSample.AlignCellsHorizontally( cellWorksheet );
        StyleSample.AlignCellsVertically( cellWorksheet );
        StyleSample.WrapTextInCell( cellWorksheet );

        // Set the width of columns.        
        cellWorksheet.Columns[ 1 ].Width = 20d;
        cellWorksheet.Columns[ 2 ].Width = 25d;
        cellWorksheet.Columns[ 3 ].Width = 25d;
        cellWorksheet.Columns[ 4 ].Width = 25d;
        cellWorksheet.Columns[ 5 ].Width = 25d;
        cellWorksheet.Columns[ 6 ].Width = 25d;
        cellWorksheet.Columns[ 7 ].Width = 13.3d;
        cellWorksheet.Columns[ 8 ].Width = 13.3d;


        // Add a second worksheet for rows.
        var rowWorksheet = workbook.Worksheets.Add( "Rows" );

        // Set the height of 6th row.        
        rowWorksheet.Rows[ 5 ].Height = 50d;

        // Set row content and alignment.
        rowWorksheet.Cells[ 5, 3 ].Value = "Setting row vertical alignment to center";
        rowWorksheet.Cells[ 5, 11 ].Value = "Another content";
        rowWorksheet.Rows[ 5 ].Style.Alignment.Vertical = VerticalAlignment.Center;


        // Add a third worksheet for column.
        var columnWorksheet = workbook.Worksheets.Add( "Columns" );

        // Set the width of 6th column.        
        columnWorksheet.Columns[ 5 ].Width = 60d;

        // Set column content and alignment.
        columnWorksheet.Cells[ 5, 5 ].Value = "Setting column horizontal alignment to center";
        columnWorksheet.Cells[ 11, 5 ].Value = "Another content";
        columnWorksheet.Columns[ 5 ].Style.Alignment.Horizontal = HorizontalAlignment.Center;

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: SetAlignments.xlsx\n" );
      }
    }

    public static void SetFonts()
    {
      using( var workbook = Workbook.Create( StyleSample.StyleSampleOutputDirectory + @"SetFonts.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var cellWorksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        cellWorksheet.Cells[ "B1" ].Value = "Set Fonts";
        cellWorksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set cell content and Font styles.
        cellWorksheet.Cells[ "C5" ].Value = "Bold, Italic and Underline";
        cellWorksheet.Cells[ "C5" ].Style.Font = new Font() { Bold = true, Italic = true, Underline = true, UnderlineType = UnderlineType.Double };

        cellWorksheet.Cells[ "C8" ].Value = "Bold and red Broadway";
        cellWorksheet.Cells[ "C8" ].Style.Font.Bold = true;
        cellWorksheet.Cells[ "C8" ].Style.Font.Color = System.Drawing.Color.Red;
        cellWorksheet.Cells[ "C8" ].Style.Font.Name = "Broadway";

        cellWorksheet.Cells[ "C12" ].Value = "Strikethrough and size 18";
        cellWorksheet.Cells[ "C12" ].Style.Font = new Font() { Strikethrough = true, Size = 18d };

        cellWorksheet.Cells[ "C15" ].Value = "Superscript";
        cellWorksheet.Cells[ "C15" ].Style.Font = new Font() { Superscript = true };

        cellWorksheet.Cells[ "D15" ].Value = "Subscript";
        cellWorksheet.Cells[ "D15" ].Style.Font = new Font() { Subscript = true };

        cellWorksheet.Cells[ "C18" ].Value = "Using theme color and tint";
        cellWorksheet.Cells[ "C18" ].Style.Font = new Font() { ThemeColor = new ThemeColor( ThemeColorType.Accent2, -0.5d ) };

        // AutoFit column "C".
        cellWorksheet.Columns[ "C" ].AutoFit();


        // Add a second worksheet for rows.
        var rowWorksheet = workbook.Worksheets.Add( "Rows" );

        // Set row content and font.
        rowWorksheet.Cells[ 5, 3 ].Value = "Setting row font to Elephant and Orange";
        rowWorksheet.Cells[ 5, 11 ].Value = "Another content";
        rowWorksheet.Rows[ 5 ].Style.Font = new Font() { Name = "Elephant", Color = System.Drawing.Color.Orange };


        // Add a third worksheet for column.
        var columnWorksheet = workbook.Worksheets.Add( "Columns" );

        // Set column content and font.
        columnWorksheet.Cells[ 5, 5 ].Value = "Setting column font to Lucida, Underline and Italic";
        columnWorksheet.Cells[ 11, 5 ].Value = "Another content";
        columnWorksheet.Columns[ 5 ].Style.Font = new Font() { Name = "Lucida Fax", Italic = true, Underline = true, UnderlineType = UnderlineType.Double };

        // AutoFit all columns with content.
        columnWorksheet.Columns.AutoFit();

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: SetFonts.xlsx\n" );
      }
    }

    public static void SetFills()
    {
      using( var workbook = Workbook.Create( StyleSample.StyleSampleOutputDirectory + @"SetFills.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var cellWorksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        cellWorksheet.Cells[ "B1" ].Value = "Set Fills";
        cellWorksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set cell content, Font Color and Fill styles.
        cellWorksheet.Cells[ "C5" ].Value = "Solid Fill";
        cellWorksheet.Cells[ "C5" ].Style.Font.Color = Color.White;
        cellWorksheet.Cells[ "C5" ].Style.Fill = new Fill() { BackgroundColor = Color.Blue };

        cellWorksheet.Cells[ "C8" ].Value = "Fill Pattern";
        cellWorksheet.Cells[ "C8" ].Style.Fill.PatternStyle = FillPattern.Gray25;
        cellWorksheet.Cells[ "C8" ].Style.Fill.PatternColor = Color.Red;
        cellWorksheet.Cells[ "C8" ].Style.Fill.BackgroundColor = Color.Yellow;

        cellWorksheet.Cells[ "C12" ].Value = "Another Fill Pattern";
        cellWorksheet.Cells[ "C12" ].Style.Fill = new Fill() { PatternStyle = FillPattern.ThinDiagonalCrosshatch, PatternColor = Color.Green };

        cellWorksheet.Cells[ "F5" ].Value = "Solid Theme Fill";
        cellWorksheet.Cells[ "F5" ].Style.Font.Color = Color.White;
        cellWorksheet.Cells[ "F5" ].Style.Fill = new Fill() { BackgroundThemeColor = new ThemeColor( ThemeColorType.Accent6 ) };

        cellWorksheet.Cells[ "F8" ].Value = "Fill Theme Pattern";
        cellWorksheet.Cells[ "F8" ].Style.Font.Color = Color.White;
        cellWorksheet.Cells[ "F8" ].Style.Fill.PatternStyle = FillPattern.DiagonalCrosshatch;
        cellWorksheet.Cells[ "F8" ].Style.Fill.PatternThemeColor = new ThemeColor( ThemeColorType.Accent5 );
        cellWorksheet.Cells[ "F8" ].Style.Fill.BackgroundThemeColor = new ThemeColor( ThemeColorType.Accent6 );

        // AutoFit column "C" to "F".
        cellWorksheet.Columns[ "C", "F" ].AutoFit();


        // Add a second worksheet for rows.
         var rowWorksheet = workbook.Worksheets.Add( "Rows" );

        // Set row content, Font color and fill.
        rowWorksheet.Cells[ 5, 3 ].Value = "Setting row fill";
        rowWorksheet.Cells[ 5, 11 ].Value = "Another content";
        rowWorksheet.Rows[ 5 ].Style.Font.Color = Color.White;
        rowWorksheet.Rows[ 5 ].Style.Fill = new Fill() { PatternStyle = FillPattern.ReverseDiagonalStripe, BackgroundColor = Color.HotPink };


        // Add a third worksheet for column.
        var columnWorksheet = workbook.Worksheets.Add( "Columns" );

        // Set column content, Font color and fill.
        columnWorksheet.Cells[ 5, 5 ].Value = "Setting column fill";
        columnWorksheet.Cells[ 11, 5 ].Value = "Another content";
        columnWorksheet.Columns[ 5 ].Style.Font.Color = Color.White;
        columnWorksheet.Columns[ 5 ].Style.Fill = new Fill() { BackgroundColor = Color.Green };

        // AutoFit all columns with content.
        columnWorksheet.Columns.AutoFit();

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: SetFills.xlsx\n" );
      }
    }

    public static void SetBorders()
    {
      using( var workbook = Workbook.Create( StyleSample.StyleSampleOutputDirectory + @"SetBorders.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var cellWorksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        cellWorksheet.Cells[ "B1" ].Value = "Set Borders";
        cellWorksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set cell content and Border styles.
        cellWorksheet.Cells[ "C5" ].Value = "Bottom";
        cellWorksheet.Cells[ "C5" ].Style.Borders[ BorderType.Bottom ].Style = LineStyle.Double;
        cellWorksheet.Cells[ "C5" ].Style.Borders[ BorderType.Bottom ].Color = Color.Red;

        cellWorksheet.Cells[ "C8" ].Value = "Top";
        cellWorksheet.Cells[ "C8" ].Style.Borders[ BorderType.Top ].Style = LineStyle.DashDot;

        cellWorksheet.Cells[ "C11" ].Value = "Right";
        cellWorksheet.Cells[ "C11" ].Style.Borders[ BorderType.Right ] = new Border() { Style = LineStyle.MediumDashed, Color = Color.Green };

        cellWorksheet.Cells[ "C14" ].Value = "Left";
        cellWorksheet.Cells[ "C14" ].Style.Borders[ BorderType.Left ] = new Border() { Style = LineStyle.DashDotDot, Color = Color.DarkSlateBlue };

        cellWorksheet.Cells[ "C17" ].Value = "Theme Bottom";
        cellWorksheet.Cells[ "C17" ].Style.Borders[ BorderType.Bottom ] = new Border() { Style = LineStyle.MediumDashDot, ThemeColor = new ThemeColor( ThemeColorType.Accent6 ) };

        cellWorksheet.Cells[ "F5" ].Value = "Diagonal Down";
        cellWorksheet.Cells[ "F5" ].Style.Borders[ BorderType.DiagonalDown ] = new Border() { Style = LineStyle.Dotted, Color = Color.DarkGoldenrod };

        cellWorksheet.Cells[ "F8" ].Value = "Diagonal Up";
        cellWorksheet.Cells[ "F8" ].Style.Borders[ BorderType.DiagonalUp ].Style = LineStyle.SlantDashDot;
        cellWorksheet.Cells[ "F8" ].Style.Borders[ BorderType.DiagonalUp ].Color = Color.DarkCyan;

        cellWorksheet.Cells[ "F11" ].Value = "Outside";
        cellWorksheet.Cells[ "F11" ].Style.Borders.SetOutline( LineStyle.Thick, Color.Blue );

        cellWorksheet.Cells[ "F14" ].Value = "Diagonals";
        cellWorksheet.Cells[ "F14" ].Style.Borders.SetDiagonals( LineStyle.Hair, Color.DeepPink );

        cellWorksheet.Cells[ "F17" ].Value = "Theme Outside";
        cellWorksheet.Cells[ "F17" ].Style.Borders.SetThemeOutline( LineStyle.Medium, new ThemeColor( ThemeColorType.Accent2 ) );

        // AutoFit column "C" and "F".
        cellWorksheet.Columns[ "C" ].AutoFit();
        cellWorksheet.Columns[ "F" ].AutoFit();


        // Add a second worksheet for rows.
        var rowWorksheet = workbook.Worksheets.Add( "Rows" );

        // Set row content and borders.
        rowWorksheet.Cells[ 5, 3 ].Value = "Setting row border";
        rowWorksheet.Cells[ 5, 11 ].Value = "Another content";
        rowWorksheet.Rows[ 5 ].Style.Borders.SetOutline( LineStyle.Thick, Color.Green );

        rowWorksheet.Cells[ 8, 4 ].Value = "Setting another row border";
        rowWorksheet.Cells[ 8, 12 ].Value = "Another content";
        rowWorksheet.Rows[ 8 ].Style.Borders[ BorderType.Bottom ].Style = LineStyle.Double;


        // Add a third worksheet for column.
        var columnWorksheet = workbook.Worksheets.Add( "Columns" );

        // Set column content and borders.
        columnWorksheet.Cells[ 5, 5 ].Value = "Setting column border";
        columnWorksheet.Cells[ 11, 5 ].Value = "Another content";
        columnWorksheet.Columns[ 5 ].Style.Borders.SetOutline( LineStyle.MediumDashDot, Color.DarkOrange );

        columnWorksheet.Cells[ 8, 8 ].Value = "Setting another column border";
        columnWorksheet.Cells[ 11, 8 ].Value = "Another content";
        columnWorksheet.Columns[ 8 ].Style.Borders[ BorderType.DiagonalDown ].Style = LineStyle.Medium;
        columnWorksheet.Columns[ 8 ].Style.Borders[ BorderType.DiagonalDown ].Color = Color.Red;

        // AutoFit all columns with content.
        columnWorksheet.Columns.AutoFit();

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: SetBorders.xlsx\n" );
      }
    }

    public static void SetFormattings()
    {
      using( var workbook = Workbook.Create( StyleSample.StyleSampleOutputDirectory + @"SetFormattings.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var cellWorksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        cellWorksheet.Cells[ "B1" ].Value = "Set Formattings";
        cellWorksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set Formatting for cells.
        StyleSample.SetFormattingWithCustomFormat( cellWorksheet );
        StyleSample.SetFormattingWithPredefinedFormatNumberId( cellWorksheet );

        // AutoFit all columns with content, from 5th row going down, and make sure column's width is between 0 and 255.
        cellWorksheet.Columns.AutoFit( 0, 255, 4 );


        // Add a second worksheet for rows.
        var rowWorksheet = workbook.Worksheets.Add( "Rows" );

        // Set row content and formatting.
        rowWorksheet.Cells[ 5, 3 ].Value = "Setting row 7 formatting to \"0.00\" using the CustomFormat:";
        rowWorksheet.Rows[ 6 ].Style.CustomFormat = "0.00";
        rowWorksheet.Cells[ 6, 4 ].Value = 50.123;
        rowWorksheet.Cells[ 6, 7 ].Value = 25;


        // Add a third worksheet for column.
        var columnWorksheet = workbook.Worksheets.Add( "Columns" );

        // Set column content and formatting.
        columnWorksheet.Cells[ 5, 1 ].Value = "Setting column D formatting to 21 (or h:mm:ss) using the PredefinedFormatNumberId:";
        columnWorksheet.Columns[ "D" ].Style.PredefinedNumberFormatId = 21;
        columnWorksheet.Cells[ "D8" ].Value = new DateTime(2022, 4, 5, 10, 5, 22 );
        columnWorksheet.Cells[ "D11" ].Value = new DateTime( 2020, 1, 1, 5, 33, 50 );



        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: SetFormattings.xlsx\n" );
      }
    }

    public static void SetBuiltinStyles()
    {
      using( var workbook = Workbook.Create( StyleSample.StyleSampleOutputDirectory + @"SetBuiltinStyles.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var cellWorksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        cellWorksheet.Cells[ "B1" ].Value = "Set Built-in Styles";
        cellWorksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set cell content and set their Built-in style types.
        cellWorksheet.Cells[ "C5" ].Value = "Good";
        cellWorksheet.Cells[ "C5" ].Style.BuiltinType = BuiltinStyleType.Good;

        cellWorksheet.Cells[ "C7" ].Value = "Bad";
        cellWorksheet.Cells[ "C7" ].Style.BuiltinType = BuiltinStyleType.Bad;

        cellWorksheet.Cells[ "C9" ].Value = "Check Cell";
        cellWorksheet.Cells[ "C9" ].Style.BuiltinType = BuiltinStyleType.CheckCell;

        cellWorksheet.Cells[ "C11" ].Value = "Accent 1";
        cellWorksheet.Cells[ "C11" ].Style.BuiltinType = BuiltinStyleType.Accent1;

        // AutoFit column "C".
        cellWorksheet.Columns[ "C" ].AutoFit();

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: SetBuiltinStyles.xlsx\n" );
      }
    }

    public static void SetStyleOnRanges()
    {
      using( var workbook = Workbook.Create( StyleSample.StyleSampleOutputDirectory + @"SetStyleOnRanges.xlsx" ) )
      {
        // Get first worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Set Style on Ranges";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Set First column style font.
        worksheet.Columns[ 0 ].Style.Font.Bold = true;

        // Set cells content.
        worksheet.Cells[ 2, 0 ].Value = "For a ColumnRange:";
        worksheet.Cells[ 3, 0 ].Value = "Horizontal Alignments on ColumnRange F to H:";
        worksheet.Cells[ 3, 5 ].Value = "One";
        worksheet.Cells[ 3, 6 ].Value = "Two";
        worksheet.Cells[ 3, 7 ].Value = "Three";

        // Get ColumnRange for columns 5 to 7(id starts at 0) and modify the alignment and Fill color of the 3 columns.
        var columnRange = worksheet.Columns[ 5, 7 ];
        columnRange.Style.Alignment.Horizontal = HorizontalAlignment.Right;
        columnRange.Style.Fill.BackgroundColor = Color.Orange;
        // Modify the width of all columns in the range.
        columnRange.Width = 15d;

        // Set cells content.
        worksheet.Cells[ 6, 0 ].Value = "For a RowRange:";
        worksheet.Cells[ 7, 0 ].Value = "Fonts on RowRange 9 to 10:";
        worksheet.Cells[ 8, 2 ].Value = "First";
        worksheet.Cells[ 8, 3 ].Value = "Second";
        worksheet.Cells[ 9, 2 ].Value = "Third";
        worksheet.Cells[ 9, 3 ].Value = "Fourth";

        // Get RowRange for rows 8 to 9(id starts at 0) and modify the fonts and fill color of the 2 rows.
        var rowRange = worksheet.Rows[ 8, 9 ];
        rowRange.Style.Font = new Font() { Color = Color.Red, Italic = true, Name = "Verdana" };
        rowRange.Style.Fill = new Fill() { PatternStyle = FillPattern.ThinDiagonalStripe, PatternColor = Color.LightGray, BackgroundColor = Color.Yellow };
        // Modify the height of all rows in the range.
        rowRange.Height = 20d;

        // Set cells content.
        worksheet.Cells[ 12, 0 ].Value = "For a CellRange:";
        worksheet.Cells[ 13, 0 ].Value = "CustomFormat on CellRange B15 to C17:";
        worksheet.Cells[ "B15" ].Value = 0.1;
        worksheet.Cells[ "C15" ].Value = 0.22;
        worksheet.Cells[ "B16" ].Value = 0.31;
        worksheet.Cells[ "C16" ].Value = 0.5;
        worksheet.Cells[ "B17" ].Value = 0.36;
        worksheet.Cells[ "C17" ].Value = 0.25;

        // Get CellRange from B15 to C17 and modify the customFormat, Fill and Borders of the 6 cells.
        var cellRange = worksheet.Cells[ "B15", "C17" ];
        cellRange.Style.CustomFormat = "0 %";
        cellRange.Style.Fill = new Fill() { PatternStyle = FillPattern.Gray25, PatternColor = Color.LightGray, BackgroundColor = Color.LightCyan };
        cellRange.Style.Borders.SetInside( LineStyle.Medium, Color.DarkGreen );
        cellRange.Style.Borders.SetOutline( LineStyle.Medium, Color.DarkGreen );

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: SetStyleOnRanges.xlsx\n" );
      }
    }

    public static void ModifyTheme()
    {
      using( var workbook = Workbook.Create( StyleSample.StyleSampleOutputDirectory + @"ModifyTheme.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var cellWorksheet = workbook.Worksheets[ 0 ];

        // Modify Workbook theme name and colors.
        workbook.Theme.Name = "My Theme";
        workbook.Theme.Colors[ ThemeColorType.Accent1 ].Color = Color.Red;
        workbook.Theme.Colors[ ThemeColorType.Text1 ].Color = Color.Green;

        // Add a title using the new Theme default Text1 color.
        cellWorksheet.Cells[ "B1" ].Value = "Modify theme";
        cellWorksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        cellWorksheet.Cells[ "C5" ].Value = "Using new Accent1 color";
        cellWorksheet.Cells[ "C5" ].Style.Font = new Font() { ThemeColor = new ThemeColor( ThemeColorType.Accent1 ) };

        cellWorksheet.Cells[ "C7" ].Value = "Using darker Accent1 color";
        cellWorksheet.Cells[ "C7" ].Style.Font = new Font() { ThemeColor = new ThemeColor( ThemeColorType.Accent1, -0.75d ) };

        cellWorksheet.Cells[ "C9" ].Value = "Using lighter Accent1 color";
        cellWorksheet.Cells[ "C9" ].Style.Font = new Font() { ThemeColor = new ThemeColor( ThemeColorType.Accent1, 0.5d ) };

        // AutoFit column "C".
        cellWorksheet.Columns[ "C" ].AutoFit();

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: ModifyTheme.xlsx\n" );
      }
    }

    public static void ChangeTextDirection()
    {
      //Some language needs to be force to right to left in order to be legible.
      using( var workbook = Workbook.Create( StyleSample.StyleSampleOutputDirectory + @"ChangeTextDirection" ) )
      {
        var worksheet = workbook.Worksheets[ 0 ];
        // Add a title 
        worksheet.Cells[ "B1" ].Value = "Change The Text Direction";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        //Some arabic text with latin chars.
        var cell = worksheet.Cells[ "B2" ];
        cell.Style.Alignment.TextDirection = TextDirectionAlignment.RightToLeft;
        cell.Value = "this ثصخقهع ففext";
        workbook.Save();
        Console.WriteLine( "\tCreated: ChangeTextDirection.xlsx\n" );
      }
    }

    public static void ChangeTextOrientation()
    {
      using( var workbook = Workbook.Create( StyleSample.StyleSampleOutputDirectory + @"ChangeTextOrientation" ) )
      {
        var worksheet = workbook.Worksheets[ 0 ];
        // Add a title 
        worksheet.Cells[ "B1" ].Value = "Change The Text Orientation";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        //Some text with a down rotation of 45 degrees ( angle is between 90 and -90 ).
        var cell = worksheet.Cells[ "B2" ];
        cell.Style.Alignment.RotationAngle = -45;
        cell.Value = "This is a rotated text.";
        workbook.Save();
        Console.WriteLine( "\tCreated: ChangeTextOrientation.xlsx\n" );
      }
    }

    #endregion

    #region Private Methods

    private static void AlignCellsHorizontally( Worksheet worksheet )
    {
      // Set cells content.
      worksheet.Cells[ 2, 0 ].Value = "Horizontal Alignments:";
      worksheet.Cells[ 2, 0 ].Style.Font = new Font() { Bold = true };
      worksheet.Cells[ 3, 1 ].Value = "General";
      worksheet.Cells[ 3, 2 ].Value = "Left";
      worksheet.Cells[ 3, 3 ].Value = "Center";
      worksheet.Cells[ 3, 4 ].Value = "Right";
      worksheet.Cells[ 3, 5 ].Value = "Fill";
      worksheet.Cells[ 3, 6 ].Value = "Center Across Selection";
      worksheet.Cells[ 3, 7 ].Value = "Justify";
      worksheet.Cells[ 3, 8 ].Value = "Distributed";

      // Set values for different types in cells. Indexing starts at 0 for rows and columns.
      worksheet.Cells[ 3, 0 ].Value = "Types";
      worksheet.Cells[ 4, 0 ].Value = "for number:";
      worksheet.Cells[ 5, 0 ].Value = "for date:";
      worksheet.Cells[ 6, 0 ].Value = "for time:";
      worksheet.Cells[ 7, 0 ].Value = "for boolean:";
      worksheet.Cells[ 8, 0 ].Value = "for text:";
      for( int i = 1; i <= 8; ++i )
      {
        worksheet.Cells[ 4, i ].Value = 225;
        worksheet.Cells[ 5, i ].Value = new DateTime( 2021, 8, 31 );
        worksheet.Cells[ 6, i ].Value = new TimeSpan( 10, 25, 0 );
        worksheet.Cells[ 7, i ].Value = true;
        worksheet.Cells[ 8, i ].Value = ( i <= 6 ) ? "A text" : "A long text showing how justification and distribution is used in a single cell.";
      }

      // Align the texts horizontally in the cells.
      for( int i = 4; i <= 8; ++i )
      {
        worksheet.Cells[ i, 1 ].Style.Alignment.Horizontal = HorizontalAlignment.General;
        worksheet.Cells[ i, 2 ].Style.Alignment.Horizontal = HorizontalAlignment.Left;
        worksheet.Cells[ i, 3 ].Style.Alignment.Horizontal = HorizontalAlignment.Center;
        worksheet.Cells[ i, 4 ].Style.Alignment.Horizontal = HorizontalAlignment.Right;
        worksheet.Cells[ i, 5 ].Style.Alignment.Horizontal = HorizontalAlignment.Fill;
        worksheet.Cells[ i, 6 ].Style.Alignment.Horizontal = HorizontalAlignment.CenterAcrossSelection;
        worksheet.Cells[ i, 7 ].Style.Alignment.Horizontal = HorizontalAlignment.Justify;
        worksheet.Cells[ i, 8 ].Style.Alignment.Horizontal = HorizontalAlignment.Distributed;
      }

      // Align center all cells from 4th row.
      worksheet.Rows[ 3 ].Style.Alignment.Horizontal = HorizontalAlignment.Center;

      // AutoFit the first column, from 4th row to 9th row, and make sure the column's width are between 0 and 255.
      worksheet.Columns[ 0 ].AutoFit( 0, 255, 3, 8 );

      // Create a table with the preceding cells.
      StyleSample.CreateFormattedTable( worksheet, 3, 0, 8, 8 );
    }

    private static void AlignCellsVertically( Worksheet worksheet )
    {
      // Set cells content.
      worksheet.Cells[ 11, 0 ].Value = "Vertical Alignments:";
      worksheet.Cells[ 11, 0 ].Style.Font = new Font() { Bold = true };
      worksheet.Cells[ 12, 1 ].Value = "Bottom";
      worksheet.Cells[ 12, 2 ].Value = "Center";
      worksheet.Cells[ 12, 3 ].Value = "Top";
      worksheet.Cells[ 12, 4 ].Value = "Justify";
      worksheet.Cells[ 12, 5 ].Value = "Distributed";

      // Set values for cells. Indexing starts at 0 for rows and columns.
      worksheet.Cells[ 12, 0 ].Value = "Types";
      worksheet.Cells[ 13, 0 ].Value = "for all types:";
      for( int i = 1; i <= 5; ++i )
      {
        worksheet.Cells[ 13, i ].Value = ( i <= 3 ) ? "A text" : "A long text showing how justification and distribution is used in a single cell.";
      }

      // Sets the height of row 14.
      worksheet.Rows[ 13 ].Height = 100;

      // Align the texts vertically in the cells.
      worksheet.Cells[ 13, 1 ].Style.Alignment.Vertical = VerticalAlignment.Bottom;
      worksheet.Cells[ 13, 2 ].Style.Alignment.Vertical = VerticalAlignment.Center;
      worksheet.Cells[ 13, 3 ].Style.Alignment.Vertical = VerticalAlignment.Top;
      worksheet.Cells[ 13, 4 ].Style.Alignment.Vertical = VerticalAlignment.Justify;
      worksheet.Cells[ 13, 5 ].Style.Alignment.Vertical = VerticalAlignment.Distributed;

      // Align center all cells from row 13.
      worksheet.Rows[ 12 ].Style.Alignment.Horizontal = HorizontalAlignment.Center;

      // Create a table with the preceding cells.
      StyleSample.CreateFormattedTable( worksheet, 12, 0, 13, 5 );
    }

    private static void WrapTextInCell( Worksheet worksheet )
    {
      // Set cells content.
      worksheet.Cells[ 16, 0 ].Value = "Wrap Text in Cell:";
      worksheet.Cells[ 16, 0 ].Style.Font = new Font() { Bold = true };
      worksheet.Cells[ 17, 1 ].Value = "This is a long text wrapping in cell B18.";
      worksheet.Cells[ 17, 3 ].Value = "This is a long text NOT wrapping in cell D18.";

      // Set Text Wrapping for the cell (17,1). Indexing starts at (0,0).
      worksheet.Cells[ 17, 1 ].Style.Alignment.IsTextWrapped = true;

      // Create a border and Background color around the preceding cells : (17, 0) to (17, 4).
      worksheet.Cells[ 17, 0, 17, 4 ].Style.Borders.SetOutline( LineStyle.Medium, Color.Black );
      worksheet.Cells[ 17, 0, 17, 4 ].Style.Fill.BackgroundColor = Color.LightBlue;
    }

    private static void CreateFormattedTable( Worksheet worksheet, int startRowId, int startColumnId, int endRowId, int endColumnId )
    {
      var table = worksheet.Tables.Add( startRowId, startColumnId, endRowId, endColumnId, TableStyle.TableStyleMedium9 );
      table.ShowFirstColumnFormatting = true;
      table.AutoFilter.ShowFilterButton = false;
    }

    private static void SetFormattingWithCustomFormat( Worksheet cellWorksheet )
    {
      cellWorksheet.Cells[ 3, 0 ].Value = "With CustomFormat:";
      cellWorksheet.Cells[ 3, 0 ].Style.Font = new Font() { Bold = true };

      // Set Cell content and formatting using the Style.CustomFormat property.
      cellWorksheet.Cells[ 4, 0 ].Value = "CustomFormat";
      cellWorksheet.Cells[ 4, 1 ].Value = "Value";
      cellWorksheet.Cells[ 4, 2 ].Value = "Result";

      cellWorksheet.Cells[ 5, 0 ].Value = "0 \"degrees\"";
      cellWorksheet.Cells[ 5, 1 ].Value = 27;
      cellWorksheet.Cells[ 5, 2 ].Value = 27;
      cellWorksheet.Cells[ 5, 2 ].Style.CustomFormat = "0 \"degrees\"";

      cellWorksheet.Cells[ 6, 0 ].Value = "0.0000";
      cellWorksheet.Cells[ 6, 1 ].Value = 33;
      cellWorksheet.Cells[ 6, 2 ].Value = 33;
      cellWorksheet.Cells[ 6, 2 ].Style.CustomFormat = "0.0000";

      cellWorksheet.Cells[ 7, 0 ].Value = "#,##0.00";
      cellWorksheet.Cells[ 7, 1 ].Value = 2024.123456;
      cellWorksheet.Cells[ 7, 2 ].Value = 2024.123456;
      cellWorksheet.Cells[ 7, 2 ].Style.CustomFormat = "#,##0.00";

      cellWorksheet.Cells[ 8, 0 ].Value = "#,##0_);[Red](#,##0)";
      cellWorksheet.Cells[ 8, 1 ].Value = -1234.56;
      cellWorksheet.Cells[ 8, 2 ].Value = -1234.56;
      cellWorksheet.Cells[ 8, 2 ].Style.CustomFormat = "#,##0_);[Red](#,##0)";

      cellWorksheet.Cells[ 9, 0 ].Value = "0%";
      cellWorksheet.Cells[ 9, 1 ].Value = 0.609;
      cellWorksheet.Cells[ 9, 2 ].Value = 0.609;
      cellWorksheet.Cells[ 9, 2 ].Style.CustomFormat = "0%";

      cellWorksheet.Cells[ 9, 0 ].Value = "m/d/yyyy";
      cellWorksheet.Cells[ 9, 1 ].Value = DateTime.Now;
      cellWorksheet.Cells[ 9, 2 ].Value = DateTime.Now;
      cellWorksheet.Cells[ 9, 2 ].Style.CustomFormat = "m/d/yyyy";

      cellWorksheet.Cells[ 10, 0 ].Value = "h:mm AM/PM";
      cellWorksheet.Cells[ 10, 1 ].Value = DateTime.Now;
      cellWorksheet.Cells[ 10, 2 ].Value = DateTime.Now;
      cellWorksheet.Cells[ 10, 2 ].Style.CustomFormat = "h:mm AM/PM";

      cellWorksheet.Cells[ 11, 0 ].Value = "mmm -yy";
      cellWorksheet.Cells[ 11, 1 ].Value = DateTime.Now;
      cellWorksheet.Cells[ 11, 2 ].Value = DateTime.Now;
      cellWorksheet.Cells[ 11, 2 ].Style.CustomFormat = "mmm -yy";

      // Create a table with the preceding cells.
      var customFormatTable = cellWorksheet.Tables.Add( 4, 0, 11, 2 );
      customFormatTable.AutoFilter.ShowFilterButton = false;
    }

    private static void SetFormattingWithPredefinedFormatNumberId( Worksheet cellWorksheet )
    {
      cellWorksheet.Cells[ 3, 5 ].Value = "With PredefinedNumberFormatId:";
      cellWorksheet.Cells[ 3, 5 ].Style.Font = new Font() { Bold = true };

      // Set Cell content and formatting using the Style.CustomFormat property.
      cellWorksheet.Cells[ 4, 5 ].Value = "PredefinedNumberFormatId";
      cellWorksheet.Cells[ 4, 6 ].Value = "Internal Format";
      cellWorksheet.Cells[ 4, 7 ].Value = "Value";
      cellWorksheet.Cells[ 4, 8 ].Value = "Result";

      cellWorksheet.Cells[ 5, 5 ].Value = "2";
      cellWorksheet.Cells[ 5, 6 ].Value = "0.00";
      cellWorksheet.Cells[ 5, 7 ].Value = 12;
      cellWorksheet.Cells[ 5, 8 ].Value = 12;
      cellWorksheet.Cells[ 5, 8 ].Style.PredefinedNumberFormatId = 2;

      cellWorksheet.Cells[ 6, 5 ].Value = "10";
      cellWorksheet.Cells[ 6, 6 ].Value = "0.00%";
      cellWorksheet.Cells[ 6, 7 ].Value = 0.25;
      cellWorksheet.Cells[ 6, 8 ].Value = 0.25;
      cellWorksheet.Cells[ 6, 8 ].Style.PredefinedNumberFormatId = 10;

      cellWorksheet.Cells[ 7, 5 ].Value = "11";
      cellWorksheet.Cells[ 7, 6 ].Value = "0.00E+00";
      cellWorksheet.Cells[ 7, 7 ].Value = 123456;
      cellWorksheet.Cells[ 7, 8 ].Value = 123465;
      cellWorksheet.Cells[ 7, 8 ].Style.PredefinedNumberFormatId = 11;

      cellWorksheet.Cells[ 8, 5 ].Value = "37";
      cellWorksheet.Cells[ 8, 6 ].Value = "#,##0 ;(#,##0)";
      cellWorksheet.Cells[ 8, 7 ].Value = 25899;
      cellWorksheet.Cells[ 8, 8 ].Value = 25899;
      cellWorksheet.Cells[ 8, 8 ].Style.PredefinedNumberFormatId = 37;

      cellWorksheet.Cells[ 9, 5 ].Value = "16";
      cellWorksheet.Cells[ 9, 6 ].Value = "d-mmm";
      cellWorksheet.Cells[ 9, 7 ].Value = DateTime.Now;
      cellWorksheet.Cells[ 9, 8 ].Value = DateTime.Now;
      cellWorksheet.Cells[ 9, 8 ].Style.PredefinedNumberFormatId = 16;

      cellWorksheet.Cells[ 10, 5 ].Value = "45";
      cellWorksheet.Cells[ 10, 6 ].Value = "mm:ss";
      cellWorksheet.Cells[ 10, 7 ].Value = DateTime.Now;
      cellWorksheet.Cells[ 10, 8 ].Value = DateTime.Now;
      cellWorksheet.Cells[ 10, 8 ].Style.PredefinedNumberFormatId = 45;

      // Create a table with the preceding cells.
      var predefinedNumberFormatIdTable = cellWorksheet.Tables.Add( 4, 5, 10, 8 );
      predefinedNumberFormatIdTable.AutoFilter.ShowFilterButton = false;
    }

    #endregion
  }
}
