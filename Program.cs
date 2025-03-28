/*using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var builder = WebApplication.CreateBuilder(args);

// Register services
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Ensure no HTTPS redirection middleware is being added.
// app.UseHttpsRedirection();  <-- This line should be removed or commented out.

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "My API V1");
    });
}

app.MapGet("/", () => "Hello, world!");

app.Run();
*/

/*using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Collections.Generic;
using System.IO;

var builder = WebApplication.CreateBuilder(args);

// Register Swagger for testing endpoints (optional)
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();


app.UseStaticFiles();

// Enable Swagger in Development
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c => 
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "File Sorting API V1");
    });
}

// Optional: a simple test endpoint
app.MapGet("/", () => "Welcome to the File Sorting Application!");

// Endpoint to trigger file sorting
app.MapPost("/sortfiles", () =>
{
    // Call our file sorting method
    var result = SortFiles();
    return Results.Ok(new { message = "The File sorting completed.", details = result });
});

app.Run();


// --- File Sorting Logic Converted from Your Core Code ---
static List<string> SortFiles()
{
    // List to collect log messages (instead of Console.WriteLine)
    var logMessages = new List<string>();

    // Specify the main folder containing all subfolders
    string mainFolder = @"C:\Users\Asus\Downloads\WebapplicationTesting";
    
    // Specify the destination folder for sorted files
    string destinationFolder = @"C:\Users\Asus\Downloads\WebapplicationTestingss";

    // Ensure the destination folder exists
    Directory.CreateDirectory(destinationFolder);

    // Get all Excel files in all subdirectories of the main folder
    string[] files = Directory.GetFiles(mainFolder, "*.xlsx", SearchOption.AllDirectories);

    // Dictionary to group files by colleague name (derived from the file name)
    var fileGroups = new Dictionary<string, List<string>>();

    foreach (string file in files)
    {
        // Get the file name without the path
        string fileName = Path.GetFileNameWithoutExtension(file);

        // Remove the date part from the file name (everything after the first underscore '_')
        string frontPart = fileName.Split('_')[0];

        // Sanitize the colleague's name (to ensure it's a valid folder name)
        string sanitizedFrontPart = SanitizeName(frontPart);

        // Add the file to the corresponding group based on the front part
        if (!fileGroups.ContainsKey(sanitizedFrontPart))
        {
            fileGroups[sanitizedFrontPart] = new List<string>();
        }
        fileGroups[sanitizedFrontPart].Add(file);
    }

    // Now, move all files for each individual to their respective folder
    foreach (var group in fileGroups)
    {
        string colleagueFolder = Path.Combine(destinationFolder, group.Key);

        // Ensure the colleague's folder exists
        Directory.CreateDirectory(colleagueFolder);

        // Move each file to the individual's folder
        foreach (var file in group.Value)
        {
            string destinationPath = Path.Combine(colleagueFolder, Path.GetFileName(file));

            // Move the file (this will throw an exception if a file already exists at the destination)
            File.Move(file, destinationPath);

            logMessages.Add($"Moved {Path.GetFileName(file)} to {colleagueFolder}");
        }
    }

    logMessages.Add("The File sorting completed successfully!");
    return logMessages;
}

// Method to sanitize file or folder names by replacing invalid characters with '_'
static string SanitizeName(string name)
{
    foreach (char invalidChar in Path.GetInvalidFileNameChars())
    {
        name = name.Replace(invalidChar, '_');
    }
    return name;
}
/*using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

var builder = WebApplication.CreateBuilder(args);

// Register Swagger for testing endpoints (optional)
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Enable static files so that files in wwwroot are served
app.UseStaticFiles();

// Enable Swagger in Development
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c => 
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "File Sorting API V1");
    });
}

// Test endpoint
app.MapGet("/", () => "Welcome to the File Sorting Application!");

// Endpoint to upload multiple files (a folder)
app.MapPost("/upload", async (HttpContext context) =>
{
    var files = context.Request.Form.Files;
    if (files.Count == 0)
        return Results.BadRequest("No files uploaded.");

    // Designated upload folder
    var uploadFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
    Directory.CreateDirectory(uploadFolder);

    foreach (var file in files)
    {
        // Optionally, you can preserve subfolder structure if needed. Here, we simply save them all in the same folder.
        var filePath = Path.Combine(uploadFolder, file.FileName);
        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }
    }

    return Results.Ok(new { message = $"{files.Count} file(s) uploaded successfully." });
});

// Endpoint to trigger file sorting on the uploaded files
app.MapPost("/sortfiles", () =>
{
    // Call the file sorting method that works on files in wwwroot/uploads
    var result = SortFiles();
    return Results.Ok(new { message = "File sorting completed.", details = result });
});

app.Run();


// --- File Sorting Logic Converted from Your Core Code ---
static List<string> SortFiles()
{
    var logMessages = new List<string>();

    // Upload folder (where files are saved)
    string mainFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");

    // Destination folder for sorted files (you can use a different folder if you prefer)
    string destinationFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "sorted");
    Directory.CreateDirectory(destinationFolder);

    // Get all Excel files in the upload folder
    string[] files = Directory.GetFiles(mainFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

    var fileGroups = new Dictionary<string, List<string>>();

    foreach (string file in files)
    {
        string fileName = Path.GetFileNameWithoutExtension(file);
        string frontPart = fileName.Split('_')[0];
        string sanitizedFrontPart = SanitizeName(frontPart);

        if (!fileGroups.ContainsKey(sanitizedFrontPart))
            fileGroups[sanitizedFrontPart] = new List<string>();

        fileGroups[sanitizedFrontPart].Add(file);
    }

    foreach (var group in fileGroups)
    {
        string colleagueFolder = Path.Combine(destinationFolder, group.Key);
        Directory.CreateDirectory(colleagueFolder);

        foreach (var file in group.Value)
        {
            string destinationPath = Path.Combine(colleagueFolder, Path.GetFileName(file));
            File.Move(file, destinationPath);
            logMessages.Add($"Moved {Path.GetFileName(file)} to {colleagueFolder}");
        }
    }

    logMessages.Add("File sorting completed successfully!");
    return logMessages;
}

static string SanitizeName(string name)
{
    foreach (char invalidChar in Path.GetInvalidFileNameChars())
    {
        name = name.Replace(invalidChar, '_');
    }
    return name;
}*/


/*using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace WebDataSortingJ
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Set EPPlus license context for non-commercial use
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var builder = WebApplication.CreateBuilder(args);

            // Register Swagger and API Explorer services
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();

            var app = builder.Build();

            // Enable static files (for serving any frontend from wwwroot)
            app.UseStaticFiles();

            // Enable Developer Exception Page and Swagger in Development
            if (app.Environment.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
                app.UseSwagger();
                app.UseSwaggerUI(c =>
                {
                    c.SwaggerEndpoint("/swagger/v1/swagger.json", "Excel Sorting API V1");
                });
            }

            // Debug endpoint to verify JSON responses
            app.MapGet("/debug", () => Results.Ok(new { message = "Debug endpoint working." }));

            // Endpoint for Dispatch Sorting
            app.MapPost("/sortdispatch", async (HttpContext context) =>
{
    try
    {
        var form = await context.Request.ReadFormAsync();
        string? inputFolder = form["inputFolder"];
        string? destinationFolder = form["destinationFolder"];

        if (string.IsNullOrWhiteSpace(inputFolder) || string.IsNullOrWhiteSpace(destinationFolder))
        {
            return Results.BadRequest(new { error = "Please provide both inputFolder and destinationFolder." });
        }

        List<string> logs = EPPlusSortDispatch(inputFolder, destinationFolder);
        return Results.Ok(new { message = "Dispatch Sorting completed.", details = logs });
    }
    catch (Exception ex)
    {
        return Results.Problem($"An error occurred: {ex.Message}");
    }
});

            // Endpoint for Scan Detail Sorting
            app.MapPost("/sortscandetail", async (HttpContext context) =>
            {
                try
                {
                    var form = await context.Request.ReadFormAsync();
                    string? inputFolder = form["inputFolder"];
                    string? destinationFolder = form["destinationFolder"];

                    if (string.IsNullOrWhiteSpace(inputFolder) || string.IsNullOrWhiteSpace(destinationFolder))
                    {
                        return Results.BadRequest(new { error = "Please provide both inputFolder and destinationFolder." });
                    }

                    List<string> logs = EPPlusSortScanDetail(inputFolder, destinationFolder);
                    if (logs == null || logs.Count == 0)
                    {
                        logs = new List<string> { "No logs available." };
                    }
                    return Results.Ok(new { message = "Scan Detail Sorting completed.", details = logs });
                }
                catch (Exception ex)
                {
                    return Results.Problem($"An error occurred in Scan Detail Sorting: {ex.Message}");
                }
            });

            // Root endpoint
            app.MapGet("/", () => Results.Ok(new { message = "Welcome to the Excel Sorting API using EPPlus!" }));

            app.Run();
        }

        // -----------------------------------------------------------------
        // EPPlus version for Dispatch Sorting
        // -----------------------------------------------------------------
        public static List<string> EPPlusSortDispatch(string inputFolder, string destinationFolder)
        {
            var logs = new List<string>();

            try
            {
                // Ensure the destination folder exists.
                Directory.CreateDirectory(destinationFolder);

                // Get all Excel files from inputFolder (including subdirectories)
                string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories);
                if (excelFiles.Length == 0)
                {
                    logs.Add("No Excel files found in the input folder.");
                    return logs;
                }

                foreach (var file in excelFiles)
                {
                    if (Path.GetFileName(file).StartsWith("~$"))
                        continue;

                    try
                    {
                        FileInfo fi = new FileInfo(file);
                        using (var package = new ExcelPackage(fi))
                        {
                            var ws = package.Workbook.Worksheets[0];
                            if (ws.Dimension == null)
                            {
                                logs.Add($"No data found in {Path.GetFileName(file)}.");
                                continue;
                            }

                            int endRow = ws.Dimension.End.Row;
                            int dataStartRow = 2; // assuming row 1 is header
                            int highlightedCount = 0;
                            var encounteredOrders = new HashSet<string>();

                            // Loop through each data row
                            for (int r = dataStartRow; r <= endRow; r++)
                            {
                                var cellB = ws.Cells[r, 2].Value;
                                var cellC = ws.Cells[r, 3].Value;
                                if (cellB == null || cellC == null)
                                    continue;

                                string waybill = cellB.ToString() ?? "";
                                string orderNumber = cellC.ToString() ?? "";

                                if (waybill.StartsWith("6666"))
                                {
                                    if (waybill.EndsWith("1"))
                                    {
                                        var targetCell = ws.Cells[r, 3];
                                        targetCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        targetCell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                        highlightedCount++;
                                    }
                                }
                                else
                                {
                                    if (!encounteredOrders.Contains(orderNumber))
                                    {
                                        var targetCell = ws.Cells[r, 3];
                                        targetCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        targetCell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                        highlightedCount++;
                                        encounteredOrders.Add(orderNumber);
                                    }
                                }
                            }

                            // Write the highlighted count two rows below the last data row
                            ws.Cells[endRow + 2, 3].Value = $"{highlightedCount} points";

                            // Save the modified workbook with a new name (_modified.xlsx)
                            string newFileName = $"{Path.GetFileNameWithoutExtension(file)}_modified.xlsx";
                            string newFilePath = Path.Combine(destinationFolder, newFileName);
                            package.SaveAs(new FileInfo(newFilePath));

                            logs.Add($"Processed {Path.GetFileName(file)}: Highlighted Count = {highlightedCount}");
                        }
                    }
                    catch (Exception ex)
                    {
                        logs.Add($"Error processing {Path.GetFileName(file)}: {ex.Message}");
                    }
                }

                logs.Add("Dispatch Sorting completed successfully using EPPlus!");
            }
            catch (Exception ex)
            {
                logs.Add("General error in EPPlusSortDispatch: " + ex.Message);
            }

            return logs;
        }

        // -----------------------------------------------------------------
        // EPPlus version for Scan Detail Sorting (sample logic; adjust as needed)
        // -----------------------------------------------------------------
        public static List<string> EPPlusSortScanDetail(string inputFolder, string destinationFolder)
        {
            var logs = new List<string>();

            try
            {
                Directory.CreateDirectory(destinationFolder);

                string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories);
                if (excelFiles.Length == 0)
                {
                    logs.Add("No Excel files found in the input folder.");
                    return logs;
                }

                foreach (var file in excelFiles)
                {
                    if (Path.GetFileName(file).StartsWith("~$"))
                        continue;

                    try
                    {
                        FileInfo fi = new FileInfo(file);
                        using (var package = new ExcelPackage(fi))
                        {
                            var ws = package.Workbook.Worksheets[0];
                            if (ws.Dimension == null)
                            {
                                logs.Add($"No data found in {Path.GetFileName(file)}.");
                                continue;
                            }

                            int endRow = ws.Dimension.End.Row;
                            int highlightedCount = 0;
                            var encounteredValues = new HashSet<string>();

                            // Sample logic: process data in column A (adjust as needed)
                            for (int r = 2; r <= endRow; r++)
                            {
                                var cell = ws.Cells[r, 1].Value;
                                if (cell == null)
                                    continue;
                                string cellText = cell.ToString() ?? "";
                                if (cellText.StartsWith("666"))
                                {
                                    if (cellText.EndsWith("1"))
                                    {
                                        ws.Cells[r, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        ws.Cells[r, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                        highlightedCount++;
                                    }
                                }
                                else
                                {
                                    if (!encounteredValues.Contains(cellText))
                                    {
                                        ws.Cells[r, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        ws.Cells[r, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                        highlightedCount++;
                                        encounteredValues.Add(cellText);
                                    }
                                }
                            }

                            ws.Cells[endRow + 2, 3].Value = $"{highlightedCount} points";
                            string newFileName = $"{Path.GetFileNameWithoutExtension(file)}_updated.xlsx";
                            string newFilePath = Path.Combine(destinationFolder, newFileName);
                            package.SaveAs(new FileInfo(newFilePath));

                            logs.Add($"Processed {Path.GetFileName(file)}: Highlighted Count = {highlightedCount}");
                        }
                    }
                    catch (Exception ex)
                    {
                        logs.Add($"Error processing {Path.GetFileName(file)}: {ex.Message}");
                    }
                }

                logs.Add("Scan Detail Sorting completed successfully using EPPlus!");
            }
            catch (Exception ex)
            {
                logs.Add("General error in EPPlusSortScanDetail: " + ex.Message);
            }

            return logs;
        }
    }
}*/

/*using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace WebDataSortingJ
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Set the EPPlus license context for non-commercial use
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var builder = WebApplication.CreateBuilder(args);
            
            // Register Swagger and API Explorer services (optional for testing)
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();

            var app = builder.Build();

            // Enable static files (if you have any frontend in wwwroot)
            app.UseStaticFiles();

            // Use Developer Exception Page and Swagger in Development
            if (app.Environment.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
                app.UseSwagger();
                app.UseSwaggerUI(c =>
                {
                    c.SwaggerEndpoint("/swagger/v1/swagger.json", "Excel Sorting API V1");
                });
            }

            // Debug endpoint to verify JSON response
            app.MapGet("/debug", () => Results.Ok(new { message = "Debug endpoint working." }));

            // Endpoint for Dispatch Sorting (processing Column C)
            app.MapPost("/sortdispatch", async (HttpContext context) =>
            {
                try
                {
                    var form = await context.Request.ReadFormAsync();
                    string? inputFolder = form["inputFolder"];
                    string? destinationFolder = form["destinationFolder"];

                    if (string.IsNullOrWhiteSpace(inputFolder) || string.IsNullOrWhiteSpace(destinationFolder))
                    {
                        return Results.BadRequest(new { error = "Please provide both inputFolder and destinationFolder." });
                    }

                    List<string> logs = EPPlusSortDispatch(inputFolder, destinationFolder);
                    return Results.Ok(new { message = "Dispatch Sorting completed.", details = logs });
                }
                catch (Exception ex)
                {
                    return Results.Problem($"An error occurred in Dispatch Sorting: {ex.Message}");
                }
            });

            // Endpoint for Scan Detail Sorting (processing Column A)
            app.MapPost("/sortscandetail", async (HttpContext context) =>
            {
                try
                {
                    var form = await context.Request.ReadFormAsync();
                    string? inputFolder = form["inputFolder"];
                    string? destinationFolder = form["destinationFolder"];

                    if (string.IsNullOrWhiteSpace(inputFolder) || string.IsNullOrWhiteSpace(destinationFolder))
                    {
                        return Results.BadRequest(new { error = "Please provide both inputFolder and destinationFolder." });
                    }

                    List<string> logs = EPPlusSortScanDetail(inputFolder, destinationFolder);
                    return Results.Ok(new { message = "Scan Detail Sorting completed.", details = logs });
                }
                catch (Exception ex)
                {
                    return Results.Problem($"An error occurred in Scan Detail Sorting: {ex.Message}");
                }
            });

            // Root endpoint
            app.MapGet("/", () => Results.Ok(new { message = "Welcome to the Excel Sorting API using EPPlus!" }));

            app.Run();
        }

        // -----------------------------------------------------------------
        // EPPlus version for Dispatch Sorting (processes Column C)
        // -----------------------------------------------------------------
        public static List<string> EPPlusSortDispatch(string inputFolder, string destinationFolder)
        {
            var logs = new List<string>();

            try
            {
                // Ensure destination folder exists
                Directory.CreateDirectory(destinationFolder);

                // Get all Excel files from inputFolder (searching in all subdirectories)
                string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories);
                if (excelFiles.Length == 0)
                {
                    logs.Add("No Excel files found in the input folder.");
                    return logs;
                }

                foreach (var file in excelFiles)
                {
                    if (Path.GetFileName(file).StartsWith("~$"))
                        continue;

                    try
                    {
                        FileInfo fi = new FileInfo(file);
                        using (var package = new ExcelPackage(fi))
                        {
                            var ws = package.Workbook.Worksheets[0];
                            if (ws.Dimension == null)
                            {
                                logs.Add($"No data found in {Path.GetFileName(file)}.");
                                continue;
                            }

                            int endRow = ws.Dimension.End.Row;
                            int dataStartRow = 2; // Assuming row 1 contains headers
                            int highlightedCount = 0;
                            var encounteredOrders = new HashSet<string>();

                            // Loop through each row of data and process Column C
                            for (int r = dataStartRow; r <= endRow; r++)
                            {
                                var cellB = ws.Cells[r, 2].Value; // Waybill number in Column B
                                var cellC = ws.Cells[r, 3].Value; // Order number in Column C
                                if (cellB == null || cellC == null)
                                    continue;

                                string waybill = cellB.ToString() ?? "";
                                string orderNumber = cellC.ToString() ?? "";

                                if (waybill.StartsWith("6666"))
                                {
                                    if (waybill.EndsWith("1"))
                                    {
                                        var targetCell = ws.Cells[r, 3]; // Process Column C
                                        targetCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        targetCell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                        highlightedCount++;
                                    }
                                }
                                else
                                {
                                    if (!encounteredOrders.Contains(orderNumber))
                                    {
                                        var targetCell = ws.Cells[r, 3]; // Process Column C
                                        targetCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        targetCell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                        highlightedCount++;
                                        encounteredOrders.Add(orderNumber);
                                    }
                                }
                            }

                            // Write the highlighted count two rows below the last row
                            ws.Cells[endRow + 2, 3].Value = $"{highlightedCount} points";

                            // Save the modified file with a new name (_modified.xlsx)
                            string newFileName = $"{Path.GetFileNameWithoutExtension(file)}_modified.xlsx";
                            string newFilePath = Path.Combine(destinationFolder, newFileName);
                            package.SaveAs(new FileInfo(newFilePath));

                            logs.Add($"Processed {Path.GetFileName(file)}: Highlighted Count = {highlightedCount}");
                        }
                    }
                    catch (Exception ex)
                    {
                        logs.Add($"Error processing {Path.GetFileName(file)}: {ex.Message}");
                    }
                }

                logs.Add("Dispatch Sorting completed successfully using EPPlus!");
            }
            catch (Exception ex)
            {
                logs.Add("General error in EPPlusSortDispatch: " + ex.Message);
            }

            return logs;
        }

        // -----------------------------------------------------------------
        // EPPlus version for Scan Detail Sorting (processes Column A)
        // -----------------------------------------------------------------
        public static List<string> EPPlusSortScanDetail(string inputFolder, string destinationFolder)
        {
            var logs = new List<string>();

            try
            {
                Directory.CreateDirectory(destinationFolder);

                string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories);
                if (excelFiles.Length == 0)
                {
                    logs.Add("No Excel files found in the input folder.");
                    return logs;
                }

                foreach (var file in excelFiles)
                {
                    if (Path.GetFileName(file).StartsWith("~$"))
                        continue;

                    try
                    {
                        FileInfo fi = new FileInfo(file);
                        using (var package = new ExcelPackage(fi))
                        {
                            var ws = package.Workbook.Worksheets[0];
                            if (ws.Dimension == null)
                            {
                                logs.Add($"No data found in {Path.GetFileName(file)}.");
                                continue;
                            }

                            int endRow = ws.Dimension.End.Row;
                            int highlightedCount = 0;
                            var encounteredValues = new HashSet<string>();

                            // Loop through each row and process Column A for Scan Detail Sorting
                            for (int r = 2; r <= endRow; r++)
                            {
                                var cell = ws.Cells[r, 1].Value; // Process Column A
                                if (cell == null)
                                    continue;
                                string cellText = cell.ToString() ?? "";
                                if (cellText.StartsWith("666"))
                                {
                                    if (cellText.EndsWith("1"))
                                    {
                                        ws.Cells[r, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        ws.Cells[r, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                                        highlightedCount++;
                                    }
                                }
                                else
                                {
                                    if (!encounteredValues.Contains(cellText))
                                    {
                                        ws.Cells[r, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        ws.Cells[r, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                                        highlightedCount++;
                                        encounteredValues.Add(cellText);
                                    }
                                }
                            }

                            ws.Cells[endRow + 2, 3].Value = $"{highlightedCount} points";
                            string newFileName = $"{Path.GetFileNameWithoutExtension(file)}_updated.xlsx";
                            string newFilePath = Path.Combine(destinationFolder, newFileName);
                            package.SaveAs(new FileInfo(newFilePath));

                            logs.Add($"Processed {Path.GetFileName(file)}: Highlighted Count = {highlightedCount}");
                        }
                    }
                    catch (Exception ex)
                    {
                        logs.Add($"Error processing {Path.GetFileName(file)}: {ex.Message}");
                    }
                }

                logs.Add("Scan Detail Sorting completed successfully using EPPlus!");
            }
            catch (Exception ex)
            {
                logs.Add("General error in EPPlusSortScanDetail: " + ex.Message);
            }

            return logs;
        }
    }
}*/

using System;
using System.Collections.Generic;
using System.IO;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.IdentityModel.Tokens;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace WebDataSortingJ
{
    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var builder = WebApplication.CreateBuilder(args);
            builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            // ✅ Fix: Ensure `jwtKey` is retrieved correctly
            //var jwtKey = builder.Configuration["Jwt:Key"] ?? throw new Exception("JWT Key is missing.");
            var jwtKey = Environment.GetEnvironmentVariable("JWT_SECRET") ?? throw new Exception("JWT Key is missing from environment variables.");
            var keyBytes = Encoding.UTF8.GetBytes(jwtKey);
            var jwtSettings = builder.Configuration.GetSection("Jwt");
            var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtSettings["Key"]));

            builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
                .AddJwtBearer(options =>
                {
                    options.TokenValidationParameters = new TokenValidationParameters
                    {
                        ValidateIssuer = true,
                        ValidateAudience = true,
                        ValidateLifetime = true,
                        ValidateIssuerSigningKey = true,
                        //ValidIssuer = builder.Configuration["Jwt:Issuer"],
                        ValidIssuer = jwtSettings["Issuer"],
                        //ValidAudience = builder.Configuration["Jwt:Audience"],
                        ValidAudience = jwtSettings["Audience"],
                        IssuerSigningKey = key  // ✅ Fix: `IssuerSigningKey` now receives `SymmetricSecurityKey`
                    };
                });

            builder.Services.AddAuthorization();
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();

            var app = builder.Build();
            app.UseStaticFiles();
            app.UseAuthentication();
            app.UseAuthorization();

            if (app.Environment.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
                app.UseSwagger();
                app.UseSwaggerUI(c =>
                {
                    c.SwaggerEndpoint("/swagger/v1/swagger.json", "Excel Sorting API V1");
                });
            }

            // 🔐 Authentication Endpoint
 app.MapPost("/login", (HttpContext context) =>
{
    if (!context.Request.HasJsonContentType())
    {
        return Results.BadRequest("Incorrect Content-Type: application/json required.");
    }
    var tokenHandler = new JwtSecurityTokenHandler();
    
    // ✅ Move these outside of the object initializer
    var jwtKey = builder.Configuration["Jwt:Key"] ?? throw new Exception("JWT Key is missing.");
    var keyBytes = Encoding.UTF8.GetBytes(jwtKey);

    var tokenDescriptor = new SecurityTokenDescriptor
    {
        Subject = new ClaimsIdentity(new[]
        {
            new Claim(ClaimTypes.Name, "user")
        }),
        Expires = DateTime.UtcNow.AddMinutes(60),
        Issuer = builder.Configuration["Jwt:Issuer"],  // ✅ Use correct config values
        Audience = builder.Configuration["Jwt:Audience"], // ✅ Use correct config values
        SigningCredentials = new SigningCredentials(
            new SymmetricSecurityKey(keyBytes),
            SecurityAlgorithms.HmacSha256Signature
        )
    };

    var token = tokenHandler.CreateToken(tokenDescriptor);
    return Results.Ok(new { token = tokenHandler.WriteToken(token) });
});

            // ✅ Debug endpoint
            app.MapGet("/debug", () => Results.Ok(new { message = "Debug endpoint working." }));

            // 🔐 Protected endpoints
app.MapPost("/sortdispatch", async (HttpContext context) =>
{
    // Ensure user is authenticated
    if (!context.User.Identity?.IsAuthenticated ?? false)
        return Results.Unauthorized();

    try
    {
        // Validate Content-Type
        if (!context.Request.HasFormContentType)
        {
            return Results.BadRequest(new { error = "Incorrect Content-Type. Expected 'multipart/form-data'." });
        }

        // Read the form data
        var form = await context.Request.ReadFormAsync();
        string inputFolder = form["inputFolder"];
        string destinationFolder = form["destinationFolder"];

        // Validate input
        if (string.IsNullOrWhiteSpace(inputFolder) || string.IsNullOrWhiteSpace(destinationFolder))
        {
            return Results.BadRequest(new { error = "Please provide both inputFolder and destinationFolder." });
        }

        // Process sorting
        List<string> logs = EPPlusSortDispatch(inputFolder, destinationFolder);
        return Results.Ok(new { message = "Dispatch Sorting completed.", details = logs });
    }
    catch (Exception ex)
    {
        return Results.Problem($"An error occurred in Dispatch Sorting: {ex.Message}");
    }
}).RequireAuthorization();
               

            app.MapPost("/sortscandetail", async (HttpContext context) =>
            {
                if (!context.User.Identity?.IsAuthenticated ?? false)
                    return Results.Unauthorized();

                try
                {
                    var form = await context.Request.ReadFormAsync();
                    string inputFolder = form["inputFolder"];
                    string destinationFolder = form["destinationFolder"];

                    if (string.IsNullOrWhiteSpace(inputFolder) || string.IsNullOrWhiteSpace(destinationFolder))
                    {
                        return Results.BadRequest(new { error = "Please provide both inputFolder and destinationFolder." });
                    }

                    List<string> logs = EPPlusSortScanDetail(inputFolder, destinationFolder);
                    return Results.Ok(new { message = "Scan Detail Sorting completed.", details = logs });
                }
                catch (Exception ex)
                {
                    return Results.Problem($"An error occurred in Scan Detail Sorting: {ex.Message}");
                }
            }).RequireAuthorization();

            // ✅ Root endpoint
            app.MapGet("/", () => Results.Ok(new { message = "Welcome to the Excel Sorting API using EPPlus!" }));

            app.Run();
        }

        // -----------------------------------------------------------------
        // ✅ EPPlus version for Dispatch Sorting (processes Column C)
        // -----------------------------------------------------------------
        public static List<string> EPPlusSortDispatch(string inputFolder, string destinationFolder)
        {
            var logs = new List<string>();

            try
            {
                Directory.CreateDirectory(destinationFolder);
                string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories);
                if (excelFiles.Length == 0)
                {
                    logs.Add("No Excel files found in the input folder.");
                    return logs;
                }

                foreach (var file in excelFiles)
                {
                    if (Path.GetFileName(file).StartsWith("~$"))
                        continue;

                    try
                    {
                        FileInfo fi = new FileInfo(file);
                        using (var package = new ExcelPackage(fi))
                        {
                            var ws = package.Workbook.Worksheets[0];
                            if (ws.Dimension == null)
                            {
                                logs.Add($"No data found in {Path.GetFileName(file)}.");
                                continue;
                            }

                            int endRow = ws.Dimension.End.Row;
                            int highlightedCount = 0;
                            var encounteredOrders = new HashSet<string>();

                            for (int r = 2; r <= endRow; r++)
                            {
                                var cellC = ws.Cells[r, 3].Value;
                                if (cellC == null) continue;

                                string orderNumber = cellC.ToString() ?? "";

                                if (!encounteredOrders.Contains(orderNumber))
                                {
                                    ws.Cells[r, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    ws.Cells[r, 3].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                    highlightedCount++;
                                    encounteredOrders.Add(orderNumber);
                                }
                            }

                            ws.Cells[endRow + 2, 3].Value = $"{highlightedCount} points";
                            string newFilePath = Path.Combine(destinationFolder, $"{Path.GetFileNameWithoutExtension(file)}_modified.xlsx");
                            package.SaveAs(new FileInfo(newFilePath));

                            logs.Add($"Processed {Path.GetFileName(file)}: Highlighted Count = {highlightedCount}");
                        }
                    }
                    catch (Exception ex)
                    {
                        logs.Add($"Error processing {Path.GetFileName(file)}: {ex.Message}");
                    }
                }

                logs.Add("Dispatch Sorting completed successfully using EPPlus!");
            }
            catch (Exception ex)
            {
                logs.Add("General error in EPPlusSortDispatch: " + ex.Message);
            }

            return logs;
        }

        // ✅ EPPlus version for Scan Detail Sorting (processes Column A)
        public static List<string> EPPlusSortScanDetail(string inputFolder, string destinationFolder)
        {
            var logs = new List<string>();

            try
            {
                Directory.CreateDirectory(destinationFolder);

                string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories);
                if (excelFiles.Length == 0)
                {
                    logs.Add("No Excel files found in the input folder.");
                    return logs;
                }

                foreach (var file in excelFiles)
                {
                    if (Path.GetFileName(file).StartsWith("~$"))
                        continue;

                    try
                    {
                        FileInfo fi = new FileInfo(file);
                        using (var package = new ExcelPackage(fi))
                        {
                            var ws = package.Workbook.Worksheets[0];
                            if (ws.Dimension == null)
                            {
                                logs.Add($"No data found in {Path.GetFileName(file)}.");
                                continue;
                            }

                            int endRow = ws.Dimension.End.Row;
                            int highlightedCount = 0;
                            var encounteredValues = new HashSet<string>();

                            // Sample logic: process data in column A (adjust as needed)
                            for (int r = 2; r <= endRow; r++)
                            {
                                var cell = ws.Cells[r, 1].Value;
                                if (cell == null)
                                    continue;
                                string cellText = cell.ToString() ?? "";
                                if (cellText.StartsWith("666"))
                                {
                                    if (cellText.EndsWith("1"))
                                    {
                                        ws.Cells[r, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        ws.Cells[r, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                        highlightedCount++;
                                    }
                                }
                                else
                                {
                                    if (!encounteredValues.Contains(cellText))
                                    {
                                        ws.Cells[r, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        ws.Cells[r, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                        highlightedCount++;
                                        encounteredValues.Add(cellText);
                                    }
                                }
                            }

                            ws.Cells[endRow + 2, 3].Value = $"{highlightedCount} points";
                            string newFileName = $"{Path.GetFileNameWithoutExtension(file)}_updated.xlsx";
                            string newFilePath = Path.Combine(destinationFolder, newFileName);
                            package.SaveAs(new FileInfo(newFilePath));

                            logs.Add($"Processed {Path.GetFileName(file)}: Highlighted Count = {highlightedCount}");
                        }
                    }
                    catch (Exception ex)
                    {
                        logs.Add($"Error processing {Path.GetFileName(file)}: {ex.Message}");
                    }
                }

                logs.Add("Scan Detail Sorting completed successfully using EPPlus!");
            }
            catch (Exception ex)
            {
                logs.Add("General error in EPPlusSortScanDetail: " + ex.Message);
            }

            return logs;
}
    }
}





