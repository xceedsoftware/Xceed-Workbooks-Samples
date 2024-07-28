using Xceed.Blazor.Workbooks.Sample.Components;

//Use a valid license key
Xceed.Workbooks.NET.Licenser.LicenseKey = "XXXXX-XXXXX-XXXXX-XXXX";

var builder = WebApplication.CreateBuilder( args );

// Add services to the container.
builder.Services.AddRazorComponents()
	.AddInteractiveServerComponents();

builder.Services.AddScoped<WorkBookCreator>();


var app = builder.Build();

// Configure the HTTP request pipeline.
if( !app.Environment.IsDevelopment() )
{
	app.UseExceptionHandler( "/Error", createScopeForErrors: true );
}

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<App>()
	.AddInteractiveServerRenderMode();

app.Run();
