using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Xceed.Blazor.Workbooks.Sample;
using Xceed.Blazor.Workbooks.Sample.Services;

//Use a valid license key
Xceed.Workbooks.NET.Licenser.LicenseKey = "LICENSE_KEY_PLACEHOLDER";

var builder = WebAssemblyHostBuilder.CreateDefault( args );
builder.RootComponents.Add<App>( "#app" );
builder.RootComponents.Add<HeadOutlet>( "head::after" );

builder.Services.AddScoped( sp => new HttpClient { BaseAddress = new Uri( builder.HostEnvironment.BaseAddress ) } );
builder.Services.AddScoped<WorkBookCreator>();

await builder.Build().RunAsync();
