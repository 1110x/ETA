using System.Reflection;

namespace ETA.Rebuild.Services;

public static class AppVersion
{
    public static string Full { get; } = ResolveVersion();

    public static string Display { get; } = "v" + Full;

    private static string ResolveVersion()
    {
        try
        {
            var asm = typeof(AppVersion).Assembly;
            var info = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            if (!string.IsNullOrWhiteSpace(info))
            {
                int plus = info.IndexOf('+');
                return plus > 0 ? info.Substring(0, plus) : info;
            }
            var name = asm.GetName().Version;
            if (name != null && name.Major + name.Minor + name.Build > 0)
                return $"{name.Major}.{name.Minor}.{name.Build}";
        }
        catch { }
        return "?";
    }
}
