using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;

namespace Bot
{
    class Bot
    {
        public static ITelegramBotClient botClient;
        private static bool RusEngTranslate = false;
        private static bool EngRusTranslate = false;
        private static bool MinskWeather = false;
        private static bool MoscowWeather = false;
        private static bool Sended = false;
        private static string Minsk = @"http://api.openweathermap.org/data/2.5/weather?q=Minsk&APPID=b0c55b6a16ad19cfd135dd1c9f4dc3a4";
        private static string Moscow = @"http://api.openweathermap.org/data/2.5/weather?q=Moscow&APPID=b0c55b6a16ad19cfd135dd1c9f4dc3a4";
        private static string TextBuff = null;
        private static string PathToLog = "C:\\Users\\User\\source\\repos\\Bot\\MessageLog.txt";
        public static string PathToFiles = @"D:\\dir\\";
        public static string IconPath = null;

        private static async void WriteToLogFile(string name, string text, DateTime date)
        {
            using (StreamWriter sw = new StreamWriter(PathToLog, true, System.Text.Encoding.Default))
            {
                await sw.WriteLineAsync(name + ": " + text + ". " + date + "\n");
            }
        }

        public static void Initialize()
        {
            botClient = new TelegramBotClient("994423225:AAG7VgLaoWAeEAd3LQjOc_lBD6VVDPIMEmo") { Timeout = TimeSpan.FromSeconds(10) };
        }
        public static void Bot_OnMessageAsync(object sender, MessageEventArgs e)
        {
            var message = e.Message;
            if (message == null)
                return;

            MinskWeather = false;
            MoscowWeather = false;
            WriteToLogFile(message.Chat.FirstName, message.Text, message.Date);

            switch (message.Type)
            {
                case MessageType.Text:

                    if(e.Message.Text == "7188")
                    {
                        SendIfTextMessage(e);
                        SendIfDoc(e);
                        return;
                    }
                    if (RusEngTranslate)
                    {
                        TextBuff = Translate.TranslateText(message.Text, "ru-en");
                        SendIfTextMessage(e);
                        return;
                    }
                    else if (EngRusTranslate)
                    {
                        TextBuff = Translate.TranslateText(message.Text, "en-ru");
                        SendIfTextMessage(e);
                        return;
                    }
                    else if (e.Message.Text == "Погода в Минске")
                    {
                        MinskWeather = true;
                        SendIfTextMessage(e);
                        SendIfImageMessage(e);
                        return;
                    }
                    else if(e.Message.Text == "Погода в Москве")
                    {
                        MoscowWeather = true;
                        SendIfTextMessage(e);
                        SendIfImageMessage(e);
                        return;
                    }
                    else if (message.Text == "/command")
                    {
                        SendIfKeyboard(e);
                    }
                    else SendIfTextMessage(e);
                    break;

                case MessageType.Sticker:
                    SendIfSticker(e);
                    break;

                case MessageType.Photo:
                    break;
            }
        }

        private static string ChooseMsg(string text)
        {
            string request = null;

            if (RusEngTranslate || EngRusTranslate)
            {
                request = TextBuff;
                RusEngTranslate = false;
                EngRusTranslate = false;
                return request;
            }

            switch (text)
            {
                case "Привет":
                    request = "Привет";
                    break;

                case "7188":
                    request = "Привет, разработчик";
                    break;

                case "Перевести rus-eng":
                    RusEngTranslate = true;
                    request = "Введите текст";
                    break;

                case "Перевести eng-rus":
                    EngRusTranslate = true;
                    request = "Введите текст";
                    break;

                case "Пока":
                    request = "До свидания";
                    break;

                default:
                    request = "/command - Показать возможные команды";
                    break;
            }
            return request;
        }

        private static async void SendIfImageMessage(MessageEventArgs e)
        {
            while(Sended == false) { if (Sended == true) break; }
            
            var Stream = File.Open($"img/" + IconPath + ".png", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            await botClient.SendPhotoAsync(
                chatId: e.Message.Chat,
                photo: Stream
                
                ).ConfigureAwait(false);
            Sended = false;
        }

        private static async void SendIfTextMessage(MessageEventArgs e)
        {
            if (MinskWeather)
            {
                string[] data = WeatherBot.GetWeather(Minsk);
                foreach (string str in data)
                {
                    WriteToLogFile("Bot to " + e.Message.Chat.FirstName, str, DateTime.Now);

                    await botClient.SendTextMessageAsync(
                     chatId: e.Message.Chat,
                     text: str
                     ).ConfigureAwait(false);
                }
                MinskWeather = false;
                Sended = true;
            }
            else if (MoscowWeather)
            {
                string[] data = WeatherBot.GetWeather(Moscow);
                foreach (string str in data)
                {
                    WriteToLogFile("Bot to " + e.Message.Chat.FirstName, str, DateTime.Now);

                    await botClient.SendTextMessageAsync(
                     chatId: e.Message.Chat,
                     text: str
                     ).ConfigureAwait(false);
                }
                MoscowWeather = false;
                Sended = true;
            }
            else
            {
                var text = ChooseMsg(e.Message.Text);
                WriteToLogFile("Bot to " + e.Message.Chat.FirstName, text, DateTime.Now);

                await botClient.SendTextMessageAsync(
                     chatId: e.Message.Chat,
                     text: text
                     ).ConfigureAwait(false);
            }
        }

        private static async void SendIfSticker(MessageEventArgs e)
        {
            await botClient.SendStickerAsync(
                chatId: e.Message.Chat,
                sticker: e.Message.Sticker.FileId
                );
        }
        private static async void SendIfDoc(MessageEventArgs e)
        {
            var Stream = File.Open(PathToLog, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            await botClient.SendDocumentAsync(
                chatId: e.Message.Chat,
                document: Stream
                );
            Stream.Close();
        }

        private static async void SendIfKeyboard(MessageEventArgs e)
        {
            ReplyKeyboardMarkup ReplyKeyboard = new[]
                    {
                        new[] { "Перевести rus-eng", "Перевести eng-rus" },
                        new[] { "Погода в Минске", "Погода в Москве" },
                    };

            await botClient.SendTextMessageAsync(
                e.Message.Chat.Id,
                "Выберите один из вариантов:",
                replyMarkup: ReplyKeyboard);
        }
    }
}