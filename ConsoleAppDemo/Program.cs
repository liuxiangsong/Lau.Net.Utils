// See https://aka.ms/new-console-template for more information
using ConsoleAppDemo;
using Lau.Net.Utils.Net;
using RestSharp;

//ConfigureService.Configure();

Console.WriteLine("Hello, World!");

using (var httpClient = new HttpClient())
{
    var response = await httpClient.GetAsync("https://localhost:7027/WeatherForecast/Stream", HttpCompletionOption.ResponseHeadersRead);

    if (response.IsSuccessStatusCode)
    {
        using (var stream = await response.Content.ReadAsStreamAsync())
        using (var reader = new System.IO.StreamReader(stream))
        {
            while (!reader.EndOfStream)
            {
                var result = await reader.ReadLineAsync();
                Console.WriteLine(result);
            }
        }
    }
    else
    {
        Console.WriteLine($"Request failed with status code: {response.StatusCode}");
    }
}