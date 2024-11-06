namespace WeatherAppOpenXML.Services;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.IO;
using System.Xml.Linq;
using WeatherAppOpenXML.Models;

public class WeatherExportService
{
    public void ExportToXml(WeatherData weatherData)
    {
        string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "WeatherData.xml");

        var weatherXml = new XElement("WeatherData",
            new XElement("Location", weatherData.Name),
            new XElement("Coordinates",
                new XElement("Latitude", weatherData.Coord.Lat),
                new XElement("Longitude", weatherData.Coord.Lon)
            ),
            new XElement("Weather",
                new XElement("Condition", weatherData.Weather.FirstOrDefault()?.Main),
                new XElement("Description", weatherData.Weather.FirstOrDefault()?.Description)
            ),
            new XElement("Temperature",
                new XElement("Current", weatherData.Main.Temp),
                new XElement("FeelsLike", weatherData.Main.Feels_like),
                new XElement("Min", weatherData.Main.Temp_min),
                new XElement("Max", weatherData.Main.Temp_max)
            ),
            new XElement("Pressure", weatherData.Main.Pressure),
            new XElement("Humidity", weatherData.Main.Humidity),
            new XElement("Wind",
                new XElement("Speed", weatherData.Wind.Speed),
                new XElement("Direction", weatherData.Wind.Deg)
            ),
            new XElement("Cloudiness", weatherData.Clouds.All),
            new XElement("Visibility", weatherData.Visibility),
            new XElement("Country", weatherData.Sys.Country),
            new XElement("Sunrise", DateTimeOffset.FromUnixTimeSeconds(weatherData.Sys.Sunrise).ToLocalTime().ToString("HH:mm")),
            new XElement("Sunset", DateTimeOffset.FromUnixTimeSeconds(weatherData.Sys.Sunset).ToLocalTime().ToString("HH:mm"))
        );

        weatherXml.Save(filePath);
    }

    public void ExportToExcel(WeatherData weatherData)
    {
        if (weatherData == null)
        {
            throw new ArgumentNullException(nameof(weatherData));
        }

        // Define the file path (you can change it to save in a specific directory)
        string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "WeatherData.xlsx");

        // Create the spreadsheet document
        using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            // Add a WorkbookPart to the document
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add a WorksheetPart to the Workbook
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook
            Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
            Sheet sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Weather Data" };
            sheets.Append(sheet);

            // Get the SheetData (where we will add rows and cells)
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // Add Header Row
            var headerRow = new Row();
            headerRow.Append(
                CreateCell("Property"),
                CreateCell("Value")
            );
            sheetData.Append(headerRow);

            // Add Data Rows
            sheetData.Append(CreateRow("Location", weatherData.Name));
            sheetData.Append(CreateRow("Latitude", weatherData.Coord.Lat.ToString()));
            sheetData.Append(CreateRow("Longitude", weatherData.Coord.Lon.ToString()));
            sheetData.Append(CreateRow("Condition", weatherData.Weather.FirstOrDefault()?.Main));
            sheetData.Append(CreateRow("Description", weatherData.Weather.FirstOrDefault()?.Description));
            sheetData.Append(CreateRow("Temperature (°C)", weatherData.Main.Temp.ToString()));
            sheetData.Append(CreateRow("Feels Like (°C)", weatherData.Main.Feels_like.ToString()));
            sheetData.Append(CreateRow("Temperature Min (°C)", weatherData.Main.Temp_min.ToString()));
            sheetData.Append(CreateRow("Temperature Max (°C)", weatherData.Main.Temp_max.ToString()));
            sheetData.Append(CreateRow("Pressure (hPa)", weatherData.Main.Pressure.ToString()));
            sheetData.Append(CreateRow("Humidity (%)", weatherData.Main.Humidity.ToString()));
            sheetData.Append(CreateRow("Wind Speed (m/s)", weatherData.Wind.Speed.ToString()));
            sheetData.Append(CreateRow("Wind Direction (°)", weatherData.Wind.Deg.ToString()));
            sheetData.Append(CreateRow("Cloudiness (%)", weatherData.Clouds.All.ToString()));
            sheetData.Append(CreateRow("Visibility (m)", weatherData.Visibility.ToString()));
            sheetData.Append(CreateRow("Country", weatherData.Sys.Country));
            sheetData.Append(CreateRow("Sunrise", DateTimeOffset.FromUnixTimeSeconds(weatherData.Sys.Sunrise).ToLocalTime().ToString("HH:mm")));
            sheetData.Append(CreateRow("Sunset", DateTimeOffset.FromUnixTimeSeconds(weatherData.Sys.Sunset).ToLocalTime().ToString("HH:mm")));

            workbookPart.Workbook.Save();
        }
    }

    // Helper method to create cells
    private Cell CreateCell(string text)
    {
        return new Cell
        {
            DataType = CellValues.String,
            CellValue = new CellValue(text)
        };
    }

    // Helper method to create rows
    private Row CreateRow(string propertyName, string value)
    {
        var row = new Row();
        row.Append(
            CreateCell(propertyName),
            CreateCell(value)
        );
        return row;
    }

    public async Task<string> ExportToExcelAsync(WeatherData weatherData)
    {
        if (weatherData == null)
        {
            throw new ArgumentNullException(nameof(weatherData));
        }

        string fileName = "WeatherData.xlsx";
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", fileName);

        using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Create a stylesheet and define custom styles
            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateStylesheet();
            stylesPart.Stylesheet.Save();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Define column widths
            Columns columns = new Columns(
                new Column { Min = 1, Max = 1, Width = 20, CustomWidth = true }, // Column 1 width (Property column)
                new Column { Min = 2, Max = 2, Width = 30, CustomWidth = true }  // Column 2 width (Value column)
            );
            worksheetPart.Worksheet.InsertAt(columns, 0);

            Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
            Sheet sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Weather Data" };
            sheets.Append(sheet);

            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // Add Header Row with Bold and Larger Font
            var headerRow = new Row();
            headerRow.Append(
                CreateStyledCell("Property", 1), // Style 1: Bold, larger font
                CreateStyledCell("Value", 1)     // Style 1: Bold, larger font
            );
            sheetData.Append(headerRow);

            // Add Data Rows with Alternating Background Colors
            bool useAlternateColor = false;
            sheetData.Append(CreateDataRow("Location", weatherData.Name, useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Latitude", weatherData.Coord.Lat.ToString(), useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Longitude", weatherData.Coord.Lon.ToString(), useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Condition", weatherData.Weather.FirstOrDefault()?.Main, useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Description", weatherData.Weather.FirstOrDefault()?.Description, useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Temperature", $"{weatherData.Main.Temp} °C (Feels like {weatherData.Main.Feels_like} °C)", useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Temperature Range", $"{weatherData.Main.Temp_min} °C to {weatherData.Main.Temp_max} °C", useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Pressure", $"{weatherData.Main.Pressure} hPa", useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Humidity", $"{weatherData.Main.Humidity}%", useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Wind", $"{weatherData.Wind.Speed} m/s at {weatherData.Wind.Deg}°", useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Cloudiness", $"{weatherData.Clouds.All}%", useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Visibility", $"{weatherData.Visibility} meters", useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Sunrise", DateTimeOffset.FromUnixTimeSeconds(weatherData.Sys.Sunrise).ToLocalTime().ToString("HH:mm"), useAlternateColor));

            useAlternateColor = !useAlternateColor;
            sheetData.Append(CreateDataRow("Sunset", DateTimeOffset.FromUnixTimeSeconds(weatherData.Sys.Sunset).ToLocalTime().ToString("HH:mm"), useAlternateColor));
           
            workbookPart.Workbook.Save();
        }
        return fileName;
    }

    private Stylesheet CreateStylesheet()
    {
        return new Stylesheet(
            new Fonts(
                new Font(), // Default font
                new Font( // Bold, larger font for title cells
                    new Bold(),
                    new FontSize() { Val = 14 })
            ),
            new Fills(
                new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue { Value = "FFFFFF" } })  { PatternType = PatternValues.Solid }), // Explicit white background fill
                new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "D3D3D3" } }) { PatternType = PatternValues.Solid }),
                new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "E0E0E0" } }) { PatternType = PatternValues.Solid })    // Lighter grey
            ),
            new Borders(new Border()), // Default border
            new CellFormats(
                new CellFormat(), // Default style
                new CellFormat { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true }, // Style 1: Bold, larger font
                new CellFormat { FontId = 0, FillId = 1, BorderId = 0, ApplyFill = true }, // Style 2: Light grey background
                new CellFormat { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true }  // Style 3: Lighter grey background
            )
        );
    }


    private Cell CreateStyledCell(string text, uint styleIndex)
    {
        return new Cell
        {
            CellValue = new CellValue(text),
            DataType = CellValues.String,
            StyleIndex = styleIndex
        };
    }

    private Row CreateDataRow(string propertyName, string value, bool useAlternateColor)
    {
        var row = new Row();
        uint styleIndex = useAlternateColor ? 0u : 3u; // Alternate between Style 0/none and Style 3, alternating Style 2 and 3 causes some dotted effect..
        row.Append(
            CreateStyledCell(propertyName, styleIndex),
            CreateStyledCell(value, styleIndex)
        );
        return row;
    }

    
}

