using SkiaSharp;
using System;
using System.IO;
using System.Linq;

namespace ETA.Services.SERVICE2;

/// <summary>
/// 생태독성 용량-반응 곡선 PNG 생성.
/// X축: log10(농도 %), Y축: 유영저해율(%), EC50 세로선 + 실측점 + 연결선.
/// 시험기록부 엑셀에 삽입할 목적의 ~640×400 이미지.
/// </summary>
public static class EcotoxicityChartGenerator
{
    /// <summary>
    /// 용량-반응 곡선 PNG 바이트 배열 반환.
    /// </summary>
    /// <param name="concentrations">농도(%) — 오름차순</param>
    /// <param name="organisms">각 농도별 생물수</param>
    /// <param name="mortalities">각 농도별 유영저해+치사수</param>
    /// <param name="ec50">EC50 (%)</param>
    /// <param name="lowerCI">95% CI 하한</param>
    /// <param name="upperCI">95% CI 상한</param>
    /// <param name="tu">독성단위 TU</param>
    /// <param name="method">"TSK" 또는 "Probit"</param>
    public static byte[] Generate(
        double[] concentrations, int[] organisms, int[] mortalities,
        double ec50, double lowerCI, double upperCI, double tu, string method)
    {
        const int W = 640, H = 400;
        const int padLeft = 70, padRight = 30, padTop = 50, padBottom = 60;
        int plotW = W - padLeft - padRight;
        int plotH = H - padTop - padBottom;

        using var surface = SKSurface.Create(new SKImageInfo(W, H));
        var canvas = surface.Canvas;
        canvas.Clear(SKColors.White);

        // ── 축 범위 (log10 스케일) ───────────────────────────────────────────
        double xMin = Math.Log10(Math.Max(concentrations.Min() / 2.0, 0.1));  // 약간 여유
        double xMax = Math.Log10(concentrations.Max() * 1.2);
        if (ec50 > 0)
        {
            xMin = Math.Min(xMin, Math.Log10(Math.Max(ec50 * 0.5, 0.1)));
            xMax = Math.Max(xMax, Math.Log10(ec50 * 1.5));
        }
        double yMin = 0, yMax = 100;

        float XtoPx(double x) => padLeft + (float)((x - xMin) / (xMax - xMin) * plotW);
        float YtoPx(double y) => padTop + plotH - (float)((y - yMin) / (yMax - yMin) * plotH);

        // ── 축/그리드 페인트 ─────────────────────────────────────────────────
        using var gridPaint = new SKPaint { Color = new SKColor(230, 230, 230), StrokeWidth = 1, IsStroke = true };
        using var axisPaint = new SKPaint { Color = SKColors.Black, StrokeWidth = 1.5f, IsStroke = true };
        using var labelPaint = new SKPaint { Color = SKColors.Black, IsAntialias = true };
        using var labelFont = new SKFont { Size = 11 };
        using var titleFont = new SKFont { Size = 14, Embolden = true };
        using var axisFont = new SKFont { Size = 12 };

        // 수평 그리드 + Y 눈금 (0, 25, 50, 75, 100%)
        foreach (var ytick in new[] { 0, 25, 50, 75, 100 })
        {
            float py = YtoPx(ytick);
            canvas.DrawLine(padLeft, py, padLeft + plotW, py, gridPaint);
            canvas.DrawText($"{ytick}", padLeft - 8, py + 4,
                SKTextAlign.Right, labelFont, labelPaint);
        }

        // 수직 그리드 + X 눈금 — 각 실제 농도 위치에 로그 눈금
        foreach (var c in concentrations)
        {
            float px = XtoPx(Math.Log10(c));
            canvas.DrawLine(px, padTop, px, padTop + plotH, gridPaint);
            canvas.DrawText($"{c:G}", px, padTop + plotH + 16,
                SKTextAlign.Center, labelFont, labelPaint);
        }

        // ── 50% 가로 기준선 (빨강 점선) ──────────────────────────────────────
        using var dashedRed = new SKPaint
        {
            Color = new SKColor(200, 50, 50), StrokeWidth = 1,
            IsStroke = true, PathEffect = SKPathEffect.CreateDash(new[] { 6f, 4f }, 0),
        };
        float y50 = YtoPx(50);
        canvas.DrawLine(padLeft, y50, padLeft + plotW, y50, dashedRed);

        // ── EC50 세로선 (파랑) ────────────────────────────────────────────────
        if (ec50 > 0)
        {
            using var dashedBlue = new SKPaint
            {
                Color = new SKColor(50, 100, 200), StrokeWidth = 1.5f,
                IsStroke = true, PathEffect = SKPathEffect.CreateDash(new[] { 5f, 3f }, 0),
            };
            float xEc = XtoPx(Math.Log10(ec50));
            canvas.DrawLine(xEc, padTop, xEc, padTop + plotH, dashedBlue);

            // 95% CI 밴드 (연파랑 반투명)
            if (lowerCI > 0 && upperCI > 0 && upperCI > lowerCI)
            {
                float xLo = XtoPx(Math.Log10(lowerCI));
                float xHi = XtoPx(Math.Log10(upperCI));
                using var ciFill = new SKPaint { Color = new SKColor(50, 100, 200, 40) };
                canvas.DrawRect(xLo, padTop, xHi - xLo, plotH, ciFill);
            }
        }

        // ── 실측점 + 연결선 ─────────────────────────────────────────────────
        using var lineP = new SKPaint { Color = new SKColor(30, 150, 60), StrokeWidth = 2, IsStroke = true, IsAntialias = true };
        using var pointP = new SKPaint { Color = new SKColor(30, 120, 50), IsAntialias = true };
        using var pointStroke = new SKPaint { Color = SKColors.White, StrokeWidth = 2, IsStroke = true, IsAntialias = true };

        var pts = new SKPoint[concentrations.Length];
        for (int i = 0; i < concentrations.Length; i++)
        {
            double pct = organisms[i] > 0 ? 100.0 * mortalities[i] / organisms[i] : 0;
            pts[i] = new SKPoint(XtoPx(Math.Log10(concentrations[i])), YtoPx(pct));
        }
        for (int i = 1; i < pts.Length; i++)
            canvas.DrawLine(pts[i - 1], pts[i], lineP);
        foreach (var p in pts)
        {
            canvas.DrawCircle(p, 5, pointP);
            canvas.DrawCircle(p, 5, pointStroke);
        }

        // ── 축 테두리 ─────────────────────────────────────────────────────────
        canvas.DrawLine(padLeft, padTop + plotH, padLeft + plotW, padTop + plotH, axisPaint); // X axis
        canvas.DrawLine(padLeft, padTop, padLeft, padTop + plotH, axisPaint);                 // Y axis

        // ── 축 제목 ───────────────────────────────────────────────────────────
        canvas.DrawText("농도 (%)", padLeft + plotW / 2f, H - 18,
            SKTextAlign.Center, axisFont, labelPaint);

        // Y축 제목 — 세로 회전
        canvas.Save();
        canvas.RotateDegrees(-90, 18, padTop + plotH / 2f);
        canvas.DrawText("유영저해율 (%)", 18, padTop + plotH / 2f + 4,
            SKTextAlign.Center, axisFont, labelPaint);
        canvas.Restore();

        // ── 제목 ──────────────────────────────────────────────────────────────
        canvas.DrawText("용량-반응 곡선 (Dose-Response)", W / 2f, 24,
            SKTextAlign.Center, titleFont, labelPaint);

        // ── 결과 요약 텍스트 박스 (우상단) ───────────────────────────────────
        if (ec50 > 0)
        {
            string line1 = $"{method}  EC50 = {ec50:F2}%";
            string line2 = $"95% CI: {lowerCI:F2} ~ {upperCI:F2}";
            string line3 = $"TU = {tu:F1}";
            using var boxBg = new SKPaint { Color = new SKColor(255, 255, 220) };
            using var boxBorder = new SKPaint { Color = SKColors.DarkGray, StrokeWidth = 1, IsStroke = true };

            float bx = padLeft + plotW - 170, by = padTop + 10;
            canvas.DrawRect(bx, by, 160, 56, boxBg);
            canvas.DrawRect(bx, by, 160, 56, boxBorder);
            canvas.DrawText(line1, bx + 8, by + 16, SKTextAlign.Left, labelFont, labelPaint);
            canvas.DrawText(line2, bx + 8, by + 32, SKTextAlign.Left, labelFont, labelPaint);
            canvas.DrawText(line3, bx + 8, by + 48, SKTextAlign.Left, labelFont, labelPaint);
        }

        // ── PNG 인코딩 ────────────────────────────────────────────────────────
        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Png, 95);
        using var ms = new MemoryStream();
        data.SaveTo(ms);
        return ms.ToArray();
    }
}
