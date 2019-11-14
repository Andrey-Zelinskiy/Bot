using Newtonsoft.Json;
using System;
using System.IO;
using System.Net;

namespace Bot
{
    // b0c55b6a16ad19cfd135dd1c9f4dc3a4
    class WeatherBot
    {
        public static string[] GetWeather(string path)
        {
            string[] gfddb = new string[6];
            string answ = "";
            OpenWeather openweather;
            WebRequest req = WebRequest.Create(path);
            req.Method = "POST";
            req.ContentType = "application/x-www-urlencoded";
            WebResponse response = req.GetResponse();
            using (Stream s = response.GetResponseStream())
            {
                using (StreamReader r = new StreamReader(s))
                {
                    answ = r.ReadToEnd();
                }
            }
            response.Close();

            openweather = JsonConvert.DeserializeObject<OpenWeather>(answ);
            gfddb[0] = "Облачность: "+openweather.clouds.all.ToString() + "%";
            gfddb[1] = "Давление: "+(openweather.main.pressure*3/4).ToString()+ " мм.рт.ст.";
            gfddb[2] = "Скорость ветра: " + openweather.wind.speed.ToString()+ " м/с";
            gfddb[3] = "Средняя температура воздуха: " + KelvinToCelsius(openweather.main.temp.ToString()) + " ℃";
            gfddb[4] = "Минимальная температура воздуха в городе: " + KelvinToCelsius(openweather.main.temp_min.ToString()) + " ℃";
            gfddb[5] = "Максимальная температура воздуха в городе: " + KelvinToCelsius(openweather.main.temp_max.ToString()) + " ℃";
            Bot.IconPath = openweather.weather[0].icon;
            return gfddb;
        }

        private static string KelvinToCelsius(string temp)
        {
            double n = Math.Round((double.Parse(temp) - 273.15), 2);
            return n.ToString();
        }
    }
}
