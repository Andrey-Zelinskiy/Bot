using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace Bot
{
    class Translate
    {
        public static string TranslateText(string s, string language)
        {
            if (s.Length > 0)//Проверка на непустую строку
            {
                //---Запрос---
                WebRequest request = WebRequest.Create("https://translate.yandex.net/api/v1.5/tr.json/translate?"
                    + "key=trnsl.1.1.20191111T145918Z.bb11707300b71e91.d3d1cdd4d512629311415751c35ba18e11c4f876"//Ключ
                    + "&text=" + s//Текст
                    + "&lang=" + language);//Язык

                WebResponse response = request.GetResponse();
                //--------------------
                //---Распарсить JSON ответ. Я скачал фреймворк Json.NET
                using (StreamReader stream = new StreamReader(response.GetResponseStream()))
                {
                    string line;
                    if ((line = stream.ReadLine()) != null)
                    {
                        Translation translation = JsonConvert.DeserializeObject<Translation>(line);
                        s = "";
                        foreach (string str in translation.text)
                        {
                            s += str;
                        }
                    }
                }
                //------------------
                return s;
            }
            else
                return "";
        }
    }

    class Translation
    {
        public string code { get; set; }
        public string lang { get; set; }
        public string[] text { get; set; }
    }
}
