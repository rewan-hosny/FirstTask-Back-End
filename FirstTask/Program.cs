using Microsoft.Extensions.FileProviders;
using OfficeOpenXml;
using Microsoft.Extensions.FileProviders;
using Microsoft.AspNetCore.Hosting;


var builder = WebApplication.CreateBuilder(args);
ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or LicenseContext.Commercial


// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddSingleton<IWebHostEnvironment>(builder.Environment);
builder.Services.AddSingleton<ExcelPackage>();



var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}
app.UseStaticFiles();

app.UseCors(builder =>
{
    builder.AllowAnyOrigin()
           .AllowAnyMethod()
           .AllowAnyHeader();
});



app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
