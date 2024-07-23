/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Annotations Sample Application
Copyright (c) 2024- Xceed Software Inc.
 
This application demonstrates how to work with Annotations when using the API 
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
  class AnnotationsSample
  {
    #region Private Members

    private const string AnnotationsSampleOutputDirectory = Program.SampleDirectory + @"Annotations\Output\";

    #endregion

    #region Constructors

    static AnnotationsSample()
    {
      if( !Directory.Exists( AnnotationsSample.AnnotationsSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( AnnotationsSample.AnnotationsSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

    public static void AddNote()
    {
      using( var workbook = Workbook.Create( AnnotationsSample.AnnotationsSampleOutputDirectory + @"AddNote.xlsx" ) )
      {
        var worksheet = workbook.Worksheets[ 0 ];
        var annotations = worksheet.Annotations;
        var formattedText = new FormattedText( "This is the title of the document" );

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Add Note";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        //Add a note without the name of the person who wrote it.
        annotations.AddNote( formattedText, "B1", false );

        formattedText.Text = " This is a new text with color";
        formattedText.Font.Color = Color.Green;

        //Add a note with the name of the person who wrote it and adding style to the text.
        annotations.AddNote( formattedText, "B3" );

        //Add a note with the name of the person and modify the text of the author name.
        annotations.AddNote( formattedText, "B5" );

        //Cast the annotation to have the notes functionalities
        Note note = (Note) worksheet.Annotations[ "B5" ];
        note[ 0 ].Font.Color = Color.Red;
        note.BackgroundColor = ColorHelper.FromIndexedColor( 11 );

        workbook.Save();
        Console.WriteLine( "\tCreated: AddNote.xlsx\n" );
      }
    }

    public static void AddComment()
    {
      using( var workbook = Workbook.Create( AnnotationsSample.AnnotationsSampleOutputDirectory + @"AddComment.xlsx" ) )
      {
        var worksheet = workbook.Worksheets[ 0 ];
        var annotations = workbook.Worksheets[ 0 ].Annotations;

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Add Comment";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        //Add a comment to the to title.
        var comment = annotations.AddComment( "This is the title of the document." , "B3" );

        //If a comment is already in a cell, it will be replied if another add is made.
        var comment2 = annotations.AddComment( "I added another comment by mistake.", "B3" );

        //Comment can also be replied by a method.
        comment2.Reply( "Wait I can also be in the thread of the conversation ?" );

        //Any comment in a thread can be used to reply in it.
        comment.Reply( "I made another mistake but looks like I will be at the end of the thread." );

        workbook.Save();
        Console.WriteLine( "\tCreated: AddComment.xlsx\n" );
      }
    }

    public static void IdentifyNotesOrComments()
    {
      using( var workbook = Workbook.Create( AnnotationsSample.AnnotationsSampleOutputDirectory + @"IdentifyNotesFromComments.xlsx" ) )
      {
        var worksheet = workbook.Worksheets[ 0 ];
        var annotations = workbook.Worksheets[ 0 ].Annotations;

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Identify Notes Or Comments";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        //Add a Comment.
        annotations.AddComment( "This is the title of the document." , "B3" );

        //Add a Note.
        annotations.AddNote( new FormattedText( "This is a note" ) , "B5" , false );

        worksheet.Cells[ "C3" ].Value = annotations[ "B3" ].AnnotationType;
        worksheet.Cells[ "C5" ].Value = annotations[ "B5" ].AnnotationType;

        workbook.Save();
        Console.WriteLine( "\tCreated: IdentifyNotesOrComments.xlsx\n" );
      }
    }

    public static void ChangeNoteFormatting()
    {
      using( var workbook = Workbook.Create( AnnotationsSample.AnnotationsSampleOutputDirectory + @"ChangeNoteFormatting.xlsx" ) )
      {
        var worksheet = workbook.Worksheets[ 0 ];
        var annotations = workbook.Worksheets[ 0 ].Annotations;

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Change Note Formatting";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

        //Adding a note will return the note.
        var note = annotations.AddNote( new FormattedText( "This is a note" ) , "B5" , false );

        //Changing the size of the note.
        note.Height = 2;
        note.Width = 3;
        note.Protection.Locked = false;
        note.Protection.LockText = false;
        note.MeasureUnit = Units.Inch;

        //Changing the text alignement.
        note.TextAlignment.Horizontal = HorizontalAlignment.Right;

        //Adding text on a new line to an existing note.
        var newFormattedText = new FormattedText( "\n Added a new independant formatted text to the existing note." );
        newFormattedText.Font.Size = 14;
        newFormattedText.Font.Italic = true;

        note.AddText( newFormattedText ); 

        workbook.Save();
        Console.WriteLine( "\tCreated: ChangeNoteFormatting.xlsx\n" );
      }
    }

     #endregion
  }
}
