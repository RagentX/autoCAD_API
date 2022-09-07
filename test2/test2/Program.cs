using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace test2
{
    public class Program
    {
        public static void onClose()
        {
            Console.WriteLine(12);
        }

        static bool onCloseEvent(int eventType)
        {
            onClose();
            return true;
        }

        static ConsoleEventDelegat handler = null;
        private delegate bool ConsoleEventDelegat(int eventType);
        [System.Runtime.InteropServices.DllImport("Kernel32.dll")]
        private static extern bool SetConsoleCtrlHandler(ConsoleEventDelegat callback, bool add);


        public static void Main(string[] args)
        {

            SetConsoleCtrlHandler(handler = new ConsoleEventDelegat(onCloseEvent), true);
            CreateHostBuilder(args).Build().Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureWebHostDefaults(webBuilder =>
                {
                    webBuilder.UseStartup<Startup>();
                });

     
    }
}
