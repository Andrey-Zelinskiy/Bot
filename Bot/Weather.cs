using System.Drawing;

namespace Bot
{
    class Weather
    {
        public int id { get; set; }
        public string main { get; set; }
        public string description { get; set; }
        public string icon { get; set; }
        public Bitmap Icon { get { return new Bitmap(Image.FromFile($"img/{icon}.png")); } }
    }
}
