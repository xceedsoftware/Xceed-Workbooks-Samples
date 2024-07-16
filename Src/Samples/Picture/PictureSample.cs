/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Picture Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to work with pictures when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/
using System;
using System.IO;

namespace Xceed.Workbooks.NET.Examples
{
  public class PictureSample
  {
    private const string PictureSampleResourcesDirectory = Program.SampleDirectory + @"Picture\Resources\";
    private const string PictureSampleOutputDirectory = Program.SampleDirectory + @"Picture\Output\";

    static PictureSample()
    {
      if( !Directory.Exists( PictureSample.PictureSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( PictureSample.PictureSampleOutputDirectory );
      }
    }

    public static void AddPicture()
    {
      using( var workbook = Workbook.Create( PictureSampleOutputDirectory + "AddPicture.xlsx") )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Add pictures"; 
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Add a picture using a stream.
        worksheet.Cells[ "A3" ].Value = "Using a stream and 2 anchors:";
        worksheet.Cells[ "A3" ].Style.Font = new Font() { Bold = true };

        var stream = new FileStream( PictureSampleResourcesDirectory + @"balloon.jpg", FileMode.Open, FileAccess.Read );
        // Add the stream picture in A4 and it should extend to E12.
        var filenamePicture = worksheet.Pictures.Add( stream, "A4", "E12" );


        // Add Picture with file name.
        worksheet.Cells[ 13, 0 ].Value = "Using a filename and 1 anchor:";
        worksheet.Cells[ 13, 0 ].Style.Font = new Font() { Bold = true };

        // Add the filename picture with its top left corner in 15th row and 1st column.
        var streamPicture = worksheet.Pictures.Add( PictureSampleResourcesDirectory + @"balloon.jpg", 14, 0 );

        workbook.Save();
        Console.WriteLine( "\tCreated: AddPicture.xlsx\n" );
      }

    }

    public static void OffsetPicture()
    {
      using( var workbook = Workbook.Create( PictureSampleOutputDirectory + "OffsetPicture.xlsx") )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Offset pictures"; 
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Add a picture using a stream.
        worksheet.Cells[ "A3" ].Value = "Original position";
        worksheet.Cells[ "A3" ].Style.Font = new Font() { Bold = true };

        var stream = new FileStream( PictureSampleResourcesDirectory + @"balloon.jpg", FileMode.Open, FileAccess.Read );
        var filenamePicture = worksheet.Pictures.Add( stream, "A4" );
        
        // Add Picture with a offset
        worksheet.Cells[ 16, 5 ].Value = "Offset position";
        worksheet.Cells[ 16, 5 ].Style.Font = new Font() { Bold = true };

        var filenamePictureOffset = worksheet.Pictures.Add( stream, "A4" );
        filenamePictureOffset.TopLeftOffsets = new Position( 3, 3, Units.Inch );

        workbook.Save();
        Console.WriteLine( "\tCreated: OffsetPicture.xlsx\n" );
      }
    }

    public static void ShrinkPictureWithOffset()
    {
      using( var workbook = Workbook.Create( PictureSampleOutputDirectory + "ShrinkPicture.xlsx") )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Shrink picture with offset"; 
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        // Add a picture using a stream.
        worksheet.Cells[ "A3" ].Value = "Original Picture";
        worksheet.Cells[ "A3" ].Style.Font = new Font() { Bold = true };
        
        var stream = new FileStream( PictureSampleResourcesDirectory + @"balloon.jpg", FileMode.Open, FileAccess.Read );
        worksheet.Pictures.Add( stream, "A4", "E12"  );
        
        // Add two anchor picture.
        worksheet.Cells[ "F3" ].Value = "Shrank Picture";
        worksheet.Cells[ "F3" ].Style.Font = new Font() { Bold = true };
        var filenamePicture = worksheet.Pictures.Add( stream, "F4", "J12" );

        //Set a negative offset to shrink the picture.
        //Warning if the offset is superior the width or the height of the picture, it will disapper.
        filenamePicture.BottomRightOffsets = new Position(-1, -1, Units.Inch);

        workbook.Save();
        Console.WriteLine( "\tCreated: ShrinkPicture.xlsx\n" );
      }
    }
  }
}
