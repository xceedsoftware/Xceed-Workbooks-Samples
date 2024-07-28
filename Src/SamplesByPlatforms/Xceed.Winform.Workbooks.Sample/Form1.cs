using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Xceed.Workbooks.NET;

namespace Xceed.Winform.Workbooks.Sample
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
			//Use a valid license key
			Xceed.Workbooks.NET.Licenser.LicenseKey = "XXXXX-XXXXX-XXXXX-XXXX";
		}

		private void SimpleWorkbook_Click( object sender, EventArgs e )
		{
			var saveFileDialog = new SaveFileDialog
			{
				Filter = "Excel Workbook|*.xlsx",
				Title = "Save an Excel Workbook"
			};

			if( saveFileDialog.ShowDialog() == DialogResult.OK )
			{
				// Create a new workbook
				var workbook = Workbook.Create( saveFileDialog.FileName );

				// Create sheets for USA, Canada, and Mexico
				var usaSheet = workbook.Worksheets[ 0 ];
				var canadaSheet = workbook.Worksheets.Add( "Canada" );
				var mexicoSheet = workbook.Worksheets.Add( "Mexico" );
				usaSheet.Name = "USA";

				// Add data to each sheet
				PopulateSheet( usaSheet, GetTopCitiesUSA() );
				PopulateSheet( canadaSheet, GetTopCitiesCanada() );
				PopulateSheet( mexicoSheet, GetTopCitiesMexico() );
				workbook.SaveAs( saveFileDialog.FileName );
				MessageBox.Show( $"Workbook saved" );
			}
		}

		private void PopulateSheet( Worksheet sheet, List<CityData> cities )
		{
			// Header
			var headers = new[] { "City", "Population", "Area (km²)", "Location (Lat, Long)" };
			for( int i = 0; i < headers.Length; i++ )
			{
				var cell = sheet.Cells[ 2, i + 2 ];
				cell.Value = headers[ i ];
				cell.Style.Font.Bold = true;
				cell.Style.Font.Size = 14;
			}

			// Data
			for( int row = 0; row < cities.Count; row++ )
			{
				sheet.Cells[ row + 3, 2 ].Value = cities[ row ].City;
				sheet.Cells[ row + 3, 3 ].Value = cities[ row ].Population;
				sheet.Cells[ row + 3, 4 ].Value = cities[ row ].Area;
				sheet.Cells[ row + 3, 5 ].Value = cities[ row ].Location;
			}
		}

		private List<CityData> GetTopCitiesUSA()
		{
			return new List<CityData>
			{
				new CityData { City = "New York", Population = "8,336,817", Area = "783.8 km²", Location = "40.7128° N, 74.0060° W" },
				new CityData { City = "Los Angeles", Population = "3,979,576", Area = "1,214.9 km²", Location = "34.0522° N, 118.2437° W" },
				new CityData { City = "Chicago", Population = "2,693,976", Area = "589.6 km²", Location = "41.8781° N, 87.6298° W" },
				new CityData { City = "Houston", Population = "2,320,268", Area = "1,651.1 km²", Location = "29.7604° N, 95.3698° W" },
				new CityData { City = "Phoenix", Population = "1,680,992", Area = "1,338.3 km²", Location = "33.4484° N, 112.0740° W" },
				new CityData { City = "Philadelphia", Population = "1,584,064", Area = "347.6 km²", Location = "39.9526° N, 75.1652° W" },
				new CityData { City = "San Antonio", Population = "1,547,253", Area = "1,208.1 km²", Location = "29.4241° N, 98.4936° W" },
				new CityData { City = "San Diego", Population = "1,423,851", Area = "964.5 km²", Location = "32.7157° N, 117.1611° W" },
				new CityData { City = "Dallas", Population = "1,343,573", Area = "881.9 km²", Location = "32.7767° N, 96.7970° W" },
				new CityData { City = "San Jose", Population = "1,021,795", Area = "469.7 km²", Location = "37.3382° N, 121.8863° W" },
				new CityData { City = "Austin", Population = "978,908", Area = "830.9 km²", Location = "30.2672° N, 97.7431° W" },
				new CityData { City = "Jacksonville", Population = "911,507", Area = "2,265.5 km²", Location = "30.3322° N, 81.6557° W" },
				new CityData { City = "Fort Worth", Population = "909,585", Area = "920.9 km²", Location = "32.7555° N, 97.3308° W" },
				new CityData { City = "Columbus", Population = "898,553", Area = "579.4 km²", Location = "39.9612° N, 82.9988° W" },
				new CityData { City = "Charlotte", Population = "885,708", Area = "771.0 km²", Location = "35.2271° N, 80.8431° W" },
				new CityData { City = "San Francisco", Population = "883,305", Area = "121.4 km²", Location = "37.7749° N, 122.4194° W" },
				new CityData { City = "Indianapolis", Population = "867,125", Area = "953.0 km²", Location = "39.7684° N, 86.1581° W" },
				new CityData { City = "Seattle", Population = "744,955", Area = "217.0 km²", Location = "47.6062° N, 122.3321° W" },
				new CityData { City = "Denver", Population = "727,211", Area = "401.3 km²", Location = "39.7392° N, 104.9903° W" },
				new CityData { City = "Washington", Population = "705,749", Area = "177.0 km²", Location = "38.9072° N, 77.0369° W" },
				new CityData { City = "Boston", Population = "692,600", Area = "125.1 km²", Location = "42.3601° N, 71.0589° W" },
				new CityData { City = "El Paso", Population = "681,728", Area = "668.6 km²", Location = "31.7619° N, 106.4850° W" },
				new CityData { City = "Detroit", Population = "670,031", Area = "370.0 km²", Location = "42.3314° N, 83.0458° W" },
				new CityData { City = "Nashville", Population = "669,053", Area = "1,362.2 km²", Location = "36.1627° N, 86.7816° W" },
				new CityData { City = "Portland", Population = "654,741", Area = "375.5 km²", Location = "45.5051° N, 122.6750° W" },
				new CityData { City = "Memphis", Population = "651,073", Area = "839.2 km²", Location = "35.1495° N, 90.0490° W" },
				new CityData { City = "Oklahoma City", Population = "649,021", Area = "1,608.4 km²", Location = "35.4676° N, 97.5164° W" },
				new CityData { City = "Las Vegas", Population = "644,644", Area = "352.0 km²", Location = "36.1699° N, 115.1398° W" },
				new CityData { City = "Louisville", Population = "617,638", Area = "1,030.0 km²", Location = "38.2527° N, 85.7585° W" },
				new CityData { City = "Baltimore", Population = "593,490", Area = "238.4 km²", Location = "39.2904° N, 76.6122° W" }
			};
		}

		private List<CityData> GetTopCitiesCanada()
		{
			return new List<CityData>
			{
				new CityData { City = "Toronto", Population = "2,731,571", Area = "630.2 km²", Location = "43.651070° N, 79.347015° W" },
				new CityData { City = "Montreal", Population = "1,704,694", Area = "431.5 km²", Location = "45.5017° N, 73.5673° W" },
				new CityData { City = "Calgary", Population = "1,239,220", Area = "825.3 km²", Location = "51.0447° N, 114.0719° W" },
				new CityData { City = "Ottawa", Population = "934,243", Area = "2,790.3 km²", Location = "45.4215° N, 75.6972° W" },
				new CityData { City = "Edmonton", Population = "932,546", Area = "684.4 km²", Location = "53.5461° N, 113.4938° W" },
				new CityData { City = "Mississauga", Population = "721,599", Area = "292.4 km²", Location = "43.5890° N, 79.6441° W" },
				new CityData { City = "Winnipeg", Population = "705,244", Area = "464.1 km²", Location = "49.8951° N, 97.1384° W" },
				new CityData { City = "Vancouver", Population = "631,486", Area = "114.7 km²", Location = "49.2827° N, 123.1207° W" },
				new CityData { City = "Brampton", Population = "593,638", Area = "266.4 km²", Location = "43.7315° N, 79.7624° W" },
				new CityData { City = "Hamilton", Population = "536,917", Area = "1,138.1 km²", Location = "43.2557° N, 79.8711° W" },
				new CityData { City = "Quebec City", Population = "531,902", Area = "454.1 km²", Location = "46.8139° N, 71.2082° W" },
				new CityData { City = "Surrey", Population = "517,887", Area = "316.4 km²", Location = "49.1044° N, 122.8011° W" },
				new CityData { City = "Laval", Population = "422,993", Area = "247.1 km²", Location = "45.6066° N, 73.7124° W" },
				new CityData { City = "Halifax", Population = "403,131", Area = "5,490.2 km²", Location = "44.6488° N, 63.5752° W" },
				new CityData { City = "London", Population = "383,822", Area = "420.5 km²", Location = "42.9849° N, 81.2453° W" },
				new CityData { City = "Markham", Population = "328,966", Area = "212.6 km²", Location = "43.8561° N, 79.3370° W" },
				new CityData { City = "Vaughan", Population = "306,233", Area = "273.5 km²", Location = "43.8361° N, 79.4983° W" },
				new CityData { City = "Gatineau", Population = "276,245", Area = "342.8 km²", Location = "45.4765° N, 75.7013° W" },
				new CityData { City = "Saskatoon", Population = "246,376", Area = "228.1 km²", Location = "52.1332° N, 106.6700° W" },
				new CityData { City = "Longueuil", Population = "239,700", Area = "115.6 km²", Location = "45.5234° N, 73.5227° W" },
				new CityData { City = "Kitchener", Population = "233,222", Area = "136.8 km²", Location = "43.4516° N, 80.4925° W" },
				new CityData { City = "Burnaby", Population = "232,755", Area = "98.6 km²", Location = "49.2488° N, 122.9805° W" },
				new CityData { City = "Windsor", Population = "217,188", Area = "146.9 km²", Location = "42.3149° N, 83.0364° W" },
				new CityData { City = "Regina", Population = "215,106", Area = "179.0 km²", Location = "50.4452° N, 104.6189° W" },
				new CityData { City = "Richmond", Population = "198,309", Area = "129.3 km²", Location = "49.1666° N, 123.1336° W" },
				new CityData { City = "Richmond Hill", Population = "195,022", Area = "101.1 km²", Location = "43.8828° N, 79.4403° W" },
				new CityData { City = "Oakville", Population = "193,832", Area = "138.5 km²", Location = "43.4675° N, 79.6877° W" },
				new CityData { City = "Burlington", Population = "183,314", Area = "185.7 km²", Location = "43.3255° N, 79.7990° W" },
				new CityData { City = "Greater Sudbury", Population = "161,531", Area = "3,228.3 km²", Location = "46.4917° N, 80.9930° W" },
				new CityData { City = "Sherbrooke", Population = "161,323", Area = "353.0 km²", Location = "45.4042° N, 71.8929° W" }
			};
		}

		private List<CityData> GetTopCitiesMexico()
		{
			return new List<CityData>
			{
				new CityData { City = "Mexico City", Population = "8,918,653", Area = "1,495 km²", Location = "19.4326° N, 99.1332° W" },
				new CityData { City = "Ecatepec", Population = "1,655,015", Area = "156.2 km²", Location = "19.6018° N, 99.0507° W" },
				new CityData { City = "Guadalajara", Population = "1,495,182", Area = "187.9 km²", Location = "20.6597° N, 103.3496° W" },
				new CityData { City = "Puebla", Population = "1,576,259", Area = "534.3 km²", Location = "19.0414° N, 98.2063° W" },
				new CityData { City = "Juarez", Population = "1,512,354", Area = "321.2 km²", Location = "31.6926° N, 106.4245° W" },
				new CityData { City = "Tijuana", Population = "1,922,523", Area = "637.5 km²", Location = "32.5149° N, 117.0382° W" },
				new CityData { City = "Leon", Population = "1,721,215", Area = "1,219.67 km²", Location = "21.1250° N, 101.6865° W" },
				new CityData { City = "Zapopan", Population = "1,476,491", Area = "893.3 km²", Location = "20.7167° N, 103.4000° W" },
				new CityData { City = "Monterrey", Population = "1,142,994", Area = "324.8 km²", Location = "25.6866° N, 100.3161° W" },
				new CityData { City = "Nezahualcoyotl", Population = "1,104,585", Area = "63.74 km²", Location = "19.4006° N, 99.0148° W" },
				new CityData { City = "Hermosillo", Population = "884,273", Area = "168.20 km²", Location = "29.0729° N, 110.9559° W" },
				new CityData { City = "Merida", Population = "892,363", Area = "858.41 km²", Location = "20.9674° N, 89.5926° W" },
				new CityData { City = "Chihuahua", Population = "925,762", Area = "247.46 km²", Location = "28.6353° N, 106.0889° W" },
				new CityData { City = "Aguascalientes", Population = "934,424", Area = "5,618 km²", Location = "21.8818° N, 102.2914° W" },
				new CityData { City = "Saltillo", Population = "807,537", Area = "3,737 km²", Location = "25.4232° N, 101.0053° W" },
				new CityData { City = "Queretaro", Population = "1,049,777", Area = "759.9 km²", Location = "20.5888° N, 100.3899° W" },
				new CityData { City = "Morelia", Population = "784,776", Area = "78.0 km²", Location = "19.7036° N, 101.1847° W" },
				new CityData { City = "Mexicali", Population = "1,029,954", Area = "13,700 km²", Location = "32.6245° N, 115.4523° W" },
				new CityData { City = "Culiacan", Population = "905,265", Area = "4,758 km²", Location = "24.8091° N, 107.3940° W" },
				new CityData { City = "Tlalnepantla", Population = "703,865", Area = "82.86 km²", Location = "19.5267° N, 99.2181° W" },
				new CityData { City = "Cancun", Population = "888,797", Area = "1,978.75 km²", Location = "21.1619° N, 86.8515° W" },
				new CityData { City = "Durango", Population = "654,876", Area = "123.9 km²", Location = "24.0277° N, 104.6532° W" },
				new CityData { City = "Reynosa", Population = "688,276", Area = "3,156.3 km²", Location = "26.0922° N, 98.2792° W" },
				new CityData { City = "San Luis Potosi", Population = "1,221,526", Area = "1,443 km²", Location = "22.1565° N, 100.9855° W" },
				new CityData { City = "Veracruz", Population = "609,964", Area = "241.2 km²", Location = "19.1738° N, 96.1342° W" },
				new CityData { City = "Villahermosa", Population = "640,359", Area = "1,612 km²", Location = "17.9869° N, 92.9303° W" },
				new CityData { City = "Tlaxcala", Population = "127,284", Area = "41.61 km²", Location = "19.3139° N, 98.2408° W" },
				new CityData { City = "Oaxaca", Population = "300,050", Area = "85.48 km²", Location = "17.0669° N, 96.7203° W" },
				new CityData { City = "Acapulco", Population = "687,608", Area = "1,882 km²", Location = "16.8531° N, 99.8237° W" },
				new CityData { City = "Tuxtla Gutierrez", Population = "598,710", Area = "412 km²", Location = "16.7528° N, 93.1156° W" }
			};
		}

		public class CityData
		{
			public string City { get; set; }
			public string Population { get; set; }
			public string Area { get; set; }
			public string Location { get; set; }
		}

		private void StyledWorkbook_Click( object sender, EventArgs e )
		{
			var saveFileDialog = new SaveFileDialog
			{
				Filter = "Excel Workbook|*.xlsx",
				Title = "Save an Excel Workbook"
			};

			if( saveFileDialog.ShowDialog() == DialogResult.OK )
			{
				// Create a new workbook
				var workbook = Workbook.Create( saveFileDialog.FileName );

				// Create sheets for each department
				var hrSheet = workbook.Worksheets[ 0 ];
				var techSheet = workbook.Worksheets.Add( "Technical Department" );
				var warehouseSheet = workbook.Worksheets.Add( "Warehouse" );

				hrSheet.Name = "Human Resources";
				// Add data to each sheet
				PopulateSheet( hrSheet, GenerateFakeData() );
				PopulateSheet( techSheet, GenerateFakeData() );
				PopulateSheet( warehouseSheet, GenerateFakeData() );
				workbook.SaveAs( saveFileDialog.FileName );
				MessageBox.Show( $"Workbook saved" );
			}
		}

		private void PopulateSheet( Worksheet sheet, List<EmployeeData> employees )
		{
			// Header
			var headers = new[] { "Name", "Surname", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "Sum", "Average" };
			for( int i = 0; i < headers.Length; i++ )
			{
				var cell = sheet.Cells[ 2, i + 2 ];
				cell.Value = headers[ i ];
				cell.Style.Font = new Xceed.Workbooks.NET.Font() { Bold = true, Color = System.Drawing.ColorTranslator.FromHtml( "#EEE4B1" ), Size = 14 };
				cell.Style.Fill.BackgroundColor = System.Drawing.ColorTranslator.FromHtml( "#430A5D" );
			}

			// Data
			for( int row = 0; row < employees.Count; row++ )
			{
				var employee = employees[ row ];
				var nameCell = sheet.Cells[ row + 3, 2 ];
				var surnameCell = sheet.Cells[ row + 3, 3 ];
				nameCell.Value = employee.Name;
				surnameCell.Value = employee.Surname;

				nameCell.Style.Fill.BackgroundColor = System.Drawing.ColorTranslator.FromHtml( "#3DC2EC" );
				nameCell.Style.Font.Italic = true;
				surnameCell.Style.Fill.BackgroundColor = System.Drawing.ColorTranslator.FromHtml( "#4B70F5" );
				surnameCell.Style.Font.Italic = true;

				for( int col = 0; col < 12; col++ )
				{
					var monthCell = sheet.Cells[ row + 3, col + 4 ];
					monthCell.Value = employee.MonthlyData[ col ];
					monthCell.Style.Fill.BackgroundColor = System.Drawing.ColorTranslator.FromHtml( "#40534C" );
					monthCell.Style.Font.Color = System.Drawing.ColorTranslator.FromHtml( "#D6BD98" );
				}

				var sumCell = sheet.Cells[ row + 3, 16 ];
				sumCell.Formula = $"SUM(D{row + 4}:O{row + 4})";
				sumCell.Style.Font.Bold = true;
				sumCell.Style.Fill.BackgroundColor = System.Drawing.ColorTranslator.FromHtml( "#481E14" );
				sumCell.Style.Font.Color = System.Drawing.ColorTranslator.FromHtml( "#F2613F" );

				var averageCell = sheet.Cells[ row + 3, 17 ];
				averageCell.Formula = $"AVERAGE(D{row + 4}:O{row + 4})";
				averageCell.Style.Font.Bold = true;
				averageCell.Style.Font.Italic = true;
				averageCell.Style.Fill.BackgroundColor = System.Drawing.ColorTranslator.FromHtml( "#32012F" );
				averageCell.Style.Font.Color = System.Drawing.ColorTranslator.FromHtml( "#F97300" );
			}
		}

		private List<EmployeeData> GenerateFakeData()
		{
			var names = new[] { "John", "Jane", "Alice", "Bob", "Carol", "David", "Eve", "Frank", "Grace", "Hank", "Ivy", "Jack", "Kathy", "Liam", "Mona" };
			var surnames = new[] { "Smith", "Johnson", "Williams", "Jones", "Brown", "Davis", "Miller", "Wilson", "Moore", "Taylor", "Anderson", "Thomas", "Jackson", "White", "Harris" };

			var random = new Random();
			var data = new List<EmployeeData>();

			for( int i = 0; i < names.Length; i++ )
			{
				var monthlyData = new int[ 12 ];
				for( int j = 0; j < 12; j++ )
				{
					monthlyData[ j ] = random.Next( 50, 150 ); // Generar datos aleatorios
				}

				data.Add( new EmployeeData
				{
					Name = names[ i ],
					Surname = surnames[ i ],
					MonthlyData = monthlyData
				} );
			}

			return data;
		}

		public class EmployeeData
		{
			public string Name { get; set; }
			public string Surname { get; set; }
			public int[] MonthlyData { get; set; }
		}
	}
}
