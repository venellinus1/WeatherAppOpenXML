using WeatherAppOpenXML.Models;

namespace WeatherAppOpenXML.Services;

public interface IWeatherService
{
    Task<WeatherData> GetWeatherDataAsync(string city);
}