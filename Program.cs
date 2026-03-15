using Avalonia;
using System;

namespace ETA;

class Program
{
    public static void Main(string[] args)
    {
        try
        {
            BuildAvaloniaApp()
                .StartWithClassicDesktopLifetime(args);
        }
        catch (Exception ex)
        {
            Console.WriteLine("🔥 프로그램 오류:");
            Console.WriteLine(ex.ToString());
            Console.ReadLine();
        }
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .LogToTrace();
}