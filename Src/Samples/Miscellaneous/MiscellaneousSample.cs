/***************************************************************************************
Xceed Workbooks for .NET – Xceed.Workbooks.NET.Examples – Miscellaneous Sample Application
Copyright (c) 2024 - Xceed Software Inc.
 
This application demonstrates how to read data from a Website and save it in Excel when using the API 
from the Xceed Workbooks for .NET.
 
This file is part of Xceed Workbooks for .NET. The source code in this file 
is only intended as a supplement to the documentation, and is provided 
"as is", without warranty of any kind, either expressed or implied.
*************************************************************************************/

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
#if NETCORE || NET5
using System.Net.Http;
#else
using System.Net;
#endif

namespace Xceed.Workbooks.NET.Examples
{
  public class MiscellaneousSample
  {
    #region Private Members

    private const string MiscellaneousSampleResourcesDirectory = Program.SampleDirectory + @"Miscellaneous\Resources\";
    private const string MiscellaneousSampleOutputDirectory = Program.SampleDirectory + @"Miscellaneous\Output\";
#if NETCORE || NET5
    static readonly HttpClient httpClient = new HttpClient();
#endif

    #endregion

    #region Constructors

    static MiscellaneousSample()
    {
      if( !Directory.Exists( MiscellaneousSample.MiscellaneousSampleOutputDirectory ) )
      {
        Directory.CreateDirectory( MiscellaneousSample.MiscellaneousSampleOutputDirectory );
      }
    }

    #endregion

    #region Public Methods

#if NETCORE || NET5
    public static async void LoadDataFromWebToExcel()
#else
    public static void LoadDataFromWebToExcel()
#endif
    {
      using( var workbook = Workbook.Create( MiscellaneousSample.MiscellaneousSampleOutputDirectory + @"LoadDataFromWebToExcel.xlsx" ) )
      {
        // Get the first worksheet. A workbook contains at least 1 worksheet.
        var worksheet = workbook.Worksheets[ 0 ];

        // Add a title.
        worksheet.Cells[ "B1" ].Value = "Load Data From Web To Excel and export to Excel";
        worksheet.Cells[ "B1" ].Style.Font = new Font() { Bold = true, Size = 15.5d };

#if NETFRAMEWORK
        // Create a request for the URL.
        var request = WebRequest.Create( "https://www.capfriendly.com/cost-per-point/2021" );

        // Get the response.
        var response = request.GetResponse();
#endif
        var dataRead = string.Empty;

        // Get the stream containing content returned by the server.
#if NETCORE || NET5
        using( var responseStream = await httpClient.GetStreamAsync( "https://www.capfriendly.com/cost-per-point/2021" ) )
#else
        using( var responseStream = response.GetResponseStream() )
#endif
        {
          // Open the stream using a StreamReader for easy access.
          var reader = new StreamReader( responseStream );
          // Read the content.
          dataRead = reader.ReadToEnd();
        }

        // Gets the rows for the data read.
        var tableRows = MiscellaneousSample.GetTableRowsFromWebSite( dataRead );

        for( int i = 0; i < tableRows.Count; ++i )
        {
          // Gets the data from a table row.
          var tableRowData = MiscellaneousSample.GetTableRowData( tableRows[ i ] );

          for( int j = 0; j < tableRowData.Count; ++j )
          {
            var currentTableRowData = tableRowData[ j ];

            // Fill the Excel workbook with each data from a row.
            if( currentTableRowData.StartsWith( "$" ) && double.TryParse( currentTableRowData.Substring( 1 ), NumberStyles.Any, CultureInfo.InvariantCulture, out double currencyValue ) )
            {
              // Set a Number and Currency format for the cell.
              worksheet.Rows[ i + 5 ].Cells[ j ].Value = currencyValue;
              worksheet.Rows[ i + 5 ].Cells[ j ].Style.CustomFormat = "$#,0";
            }
            else if( double.TryParse( currentTableRowData, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleValue ) )
            {
              // Set a Number for the cell.
              worksheet.Rows[ i + 5 ].Cells[ j ].Value = doubleValue;
            }
            else
            {
              // Set a string for the cell.
              worksheet.Rows[ i + 5 ].Cells[ j ].Value = currentTableRowData;
            }
          }
        }

        // Adjust columns width.
        worksheet.Columns[ 1 ].Width = 17;

        // Save workbook to disk.
        workbook.Save();
        Console.WriteLine( "\tCreated: LoadDataFromWebToExcel.xlsx\n" );
      }
    }

    #endregion

    #region Private Methods

    private static List<string> GetTableRowsFromWebSite( string responseFromServer )
    {
      var tableRows = new List<string>();

      var tableStartIndex = responseFromServer.IndexOf( "<table" );
      var tableEndIndex = responseFromServer.IndexOf( "</table>" );

      int counter = tableStartIndex;
      while( counter < tableEndIndex )
      {
        var tableRowStartIndex = responseFromServer.IndexOf( "<tr", counter );
        if( tableRowStartIndex < 0 )
          break;
        var tableRowEndIndex = responseFromServer.IndexOf( "</tr>", tableRowStartIndex + 3 );
        if( tableRowEndIndex < 0 )
          break;

        var tableRow = responseFromServer.Substring( tableRowStartIndex, tableRowEndIndex - tableRowStartIndex );
        if( !string.IsNullOrEmpty( tableRow ) )
        {
          tableRows.Add( tableRow );
        }

        counter = tableRowEndIndex + 5;
      }

      return tableRows;
    }

    private static List<string> GetTableRowData( string row )
    {
      var tableRowData = new List<string>();

      var counter = 0;
      while( counter < row.Length )
      {
        // Get table header.
        var thStartIndex = row.IndexOf( "<th", counter );
        if( thStartIndex >= 0 )
        {
          var thEndIndex = row.IndexOf( "</th>", thStartIndex );
          if( thEndIndex < 0 )
            break;

          var nextThStartTag = row.IndexOf( "<th", thStartIndex + 3 );
          if( nextThStartTag >= 0 )
          {
            thEndIndex = Math.Min( thEndIndex, nextThStartTag );
          }

          var thTag = row.Substring( thStartIndex, thEndIndex + 1 - thStartIndex );
          // Ignore hidden table header.
          if( thTag.Contains( "hide" ) )
          {
            counter = thEndIndex;
          }
          else
          {
            // Get content from table header.
            var content = MiscellaneousSample.GetTagContent( thTag );
            if( !string.IsNullOrEmpty( content ) )
            {
              tableRowData.Add( content );
            }
            counter = thEndIndex;
          }
        }
        else
        {
          // Get table data
          var tdStartIndex = row.IndexOf( "<td", counter );
          if( tdStartIndex < 0 )
            break;

          var tdEndIndex = row.IndexOf( "</td>", tdStartIndex );
          if( tdEndIndex < 0 )
            break;

          var tdTag = row.Substring( tdStartIndex, tdEndIndex + 5 - tdStartIndex );
          // Get content from table data.
          var content = MiscellaneousSample.GetTagContent( tdTag );
          if( !string.IsNullOrEmpty( content ) )
          {
            tableRowData.Add( content );
          }
          counter = tdEndIndex + 5;
        }
      }

      return tableRowData;
    }

    private static string GetTagContent( string tag )
    {
      var counter = 0;
      while( counter < tag.Length )
      {
        var closingTagIndex = tag.IndexOf( ">", counter );
        if( ( closingTagIndex < 0 ) || ( closingTagIndex == tag.Length - 1 ) )
          break;

        // Found content.
        if( tag[ closingTagIndex + 1 ] != '<' )
        {
          var dataIndex = closingTagIndex + 1;
          var openingTagIndex = tag.IndexOf( "<", dataIndex );

          return tag.Substring( dataIndex, openingTagIndex - dataIndex );
        }
        // Found another tag, continue searching for content.
        else
        {
          counter = closingTagIndex + 1;
        }
      }

      return string.Empty;
    }

    #endregion
  }
}
