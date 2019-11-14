using System;
using Telegram.Bot;
using Telegram.Bot.Args;

namespace Bot
{
    class Program
    {
        static void Main(string[] args)
        {
            Bot.Initialize();
            Bot.botClient.OnMessage += Bot.Bot_OnMessageAsync;
            Bot.botClient.StartReceiving();
            Console.ReadKey();
        }

        
    }
}
