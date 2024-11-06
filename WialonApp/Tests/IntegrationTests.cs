namespace WeatherAppOpenXML.Tests;

using NUnit.Framework;
using Moq;
using System.Threading.Tasks;
using WeatherAppOpenXML.Services;
using WeatherAppOpenXML.Models;


[TestFixture]
public class IntegrationTests
{
    private Mock<IWeatherService> mockWeatherService;
    private WeatherExportService weatherExportService;

    [SetUp]
    public void Setup()
    {
        weatherExportService = new WeatherExportService();
    }    

    [Test]
    public async Task ExportToExcel_CreatesFile_WhenDataIsValid()
    { 
        // Arrange 
        var weatherData = new WeatherData
        {
            Coord = new Coord { Lat = 43.417, Lon = 24.606 },
            Weather = new List<Weather> { new Weather { Main = "Clear", Description = "clear sky" } },
            Main = new Main { Temp = 15.5, Feels_like = 14.5, Temp_min = 14.0, Temp_max = 16.0, Pressure = 1012, Humidity = 60 },
            Wind = new Wind { Speed = 3.1, Deg = 120 },
            Clouds = new Clouds { All = 0 },
            Sys = new Sys { Country = "BG", Sunrise = 1600000000, Sunset = 1600040000 },
            Visibility = 10000,
            Timezone = 7200,
            Name = "Pleven"
        };
        var expectedFileName = "WeatherData.xlsx";
        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", expectedFileName);

        Directory.CreateDirectory(Path.Combine(Directory.GetCurrentDirectory(), "wwwroot"));


        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }

        // Act 
        var result = await weatherExportService.ExportToExcelAsync(weatherData);

        // Assert
        Assert.AreEqual(expectedFileName, Path.GetFileName(result), "The filename should match the expected name.");

    }
}
