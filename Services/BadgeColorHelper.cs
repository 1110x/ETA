namespace ETA.Services;

/// <summary>약칭 첫 글자 초성에 따른 배지 색상 반환</summary>
public static class BadgeColorHelper
{
    public static (string Bg, string Fg) GetBadgeColor(string 약칭)
    {
        if (string.IsNullOrEmpty(약칭)) return ("#2a2a2a", "#aaaaaa");
        char c = 약칭[0];
        if (c < '가' || c > '힣') return ("#2a2a2a", "#aaaaaa");
        int cho = (c - 0xAC00) / (21 * 28);
        return cho switch
        {
            0  => ("#1a3a1a", "#88cc88"),  // ㄱ
            1  => ("#1a2a3a", "#88aacc"),  // ㄲ
            2  => ("#2a1a3a", "#aa88cc"),  // ㄴ
            3  => ("#3a2a1a", "#ccaa88"),  // ㄷ
            4  => ("#1a3a3a", "#88ccbb"),  // ㄸ
            5  => ("#2a1a2a", "#cc88aa"),  // ㄹ
            6  => ("#1a2a2a", "#88cccc"),  // ㅁ
            7  => ("#1a1a3a", "#8888cc"),  // ㅂ
            8  => ("#2a3a1a", "#aacc88"),  // ㅃ
            9  => ("#3a1a1a", "#cc8888"),  // ㅅ
            10 => ("#2a3a3a", "#88ccee"),  // ㅆ
            11 => ("#2a2a1a", "#cccc88"),  // ㅇ
            12 => ("#1a3a2a", "#88ccaa"),  // ㅈ
            13 => ("#3a3a1a", "#ccccaa"),  // ㅉ
            14 => ("#3a1a2a", "#cc88bb"),  // ㅊ
            15 => ("#2a2a3a", "#aaaacc"),  // ㅋ
            16 => ("#1a2a1a", "#88cc99"),  // ㅌ
            17 => ("#3a1a3a", "#cc88cc"),  // ㅍ
            18 => ("#3a2a3a", "#ccaacc"),  // ㅎ
            _  => ("#2a2a2a", "#aaaaaa"),
        };
    }
}
