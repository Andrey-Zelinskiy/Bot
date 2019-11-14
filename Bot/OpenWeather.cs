using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace Bot
{
    class OpenWeather
    {
        public Coord coord { get; set; }
        public Weather[] weather { get; set; }

        [JsonProperty("base")] //так как слово base является ключевым для C# необходимо воспользоваться параметром, он как бы подменяет имя.
        public string base1 { get; set; }
        public Main main { get; set; }
        public int visibility { get; set; }
        public Wind wind { get; set; }
        public Clouds clouds { get; set; }
        public long dt { get; set; }
        public Sys sys { get; set; }
        public int id { get; set; }
        public string name { get; set; }
        public int cod { get; set; }
    }
}
