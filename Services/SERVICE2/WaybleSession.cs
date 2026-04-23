using System;
using System.Net;
using System.Net.Http;

namespace ETA.Services.SERVICE2;

/// <summary>
/// Wayble HTTP 세션 싱글톤. CDP 로그인에서 획득한 stpsession 쿠키를 seed 하고,
/// 이후 업로드/조회 호출 시 이 HttpClient 를 재사용한다.
/// </summary>
public static class WaybleSession
{
    private static readonly Uri RewaterUri = new("https://rewater.wayble.eco");

    private static CookieContainer _cookies = new();
    private static HttpClient?     _http;
    private static string?         _sessionValue;

    public static bool    Connected    => !string.IsNullOrEmpty(_sessionValue);
    public static string? SessionCookie => _sessionValue;

    public static HttpClient Client
    {
        get
        {
            if (_http != null) return _http;
            _http = BuildClient(_cookies);
            return _http;
        }
    }

    public static void SeedStpSession(string value)
    {
        _sessionValue = value;
        _cookies.Add(RewaterUri, new Cookie("stpsession", value)
        {
            Domain = "rewater.wayble.eco",
            Path   = "/",
        });
    }

    public static void Clear()
    {
        _sessionValue = null;
        _cookies      = new CookieContainer();
        _http?.Dispose();
        _http = null;
    }

    private static HttpClient BuildClient(CookieContainer jar)
    {
        var handler = new HttpClientHandler
        {
            CookieContainer = jar,
            UseCookies      = true,
        };
        var client = new HttpClient(handler)
        {
            Timeout = TimeSpan.FromSeconds(30),
        };
        client.DefaultRequestHeaders.UserAgent.ParseAdd(
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 " +
            "(KHTML, like Gecko) Chrome/122.0 Safari/537.36");
        return client;
    }
}
