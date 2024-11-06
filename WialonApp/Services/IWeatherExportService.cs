using WeatherAppOpenXML.Models;

namespace WeatherAppOpenXML.Services;

public interface IWeatherExportService
{
    Task<string> ExportToExcelAsync(WeatherData weatherData);
}