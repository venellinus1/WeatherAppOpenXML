namespace WeatherAppOpenXML.Services;
using WeatherAppOpenXML.Models;


public class WeatherService
{
    private readonly HttpClient _httpClient;
    private readonly string _apiKey = "1a2b9a7cb84b98711eb0d7aa9865ca50";
    
    public WeatherService(HttpClient httpClient)
    {
        _httpClient = httpClient;
    }

    public async Task<WeatherData> GetWeatherDataAsync(string city)
    {
        var response = await _httpClient.GetFromJsonAsync<WeatherData>($"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={_apiKey}&units=metric");
        return response;
    }
}

