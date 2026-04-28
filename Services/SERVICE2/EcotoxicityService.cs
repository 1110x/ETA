using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Services.SERVICE2;

/// <summary>
/// 생태독성 통계 분석 서비스 (ES 04704.1c 물벼룩 급성독성시험법 준수)
/// - TSK (Trimmed Spearman-Karber) + Probit
/// - 표준 농도: 100%, 50%, 25%, 12.5%, 6.25%
/// - 반복구: 5마리 × 4반복 = 20마리/농도
/// - 시험시간: 24시간
/// - 결과: EC50(%) → TU = 100/EC50
/// </summary>
public static class EcotoxicityService
{
    /// <summary>표준 시험 농도 (%, 오름차순)</summary>
    public static readonly double[] StandardConcentrations = { 6.25, 12.5, 25.0, 50.0, 100.0 };

    /// <summary>표준 시험 생물수 (5마리 × 4반복)</summary>
    public const int StandardOrganismsPerConc = 20;

    /// <summary>대조군 허용 사망률 상한 (10% 초과 시 시험 무효)</summary>
    public const double MaxControlMortality = 0.10;

    public sealed record EcotoxResult(
        double EC50, double LowerCI, double UpperCI,
        double TU, string Method,
        double TrimPercent, bool Smoothed,
        string? Warning = null);

    /// <summary>
    /// EC50 산출 불가 시 TU 계산 (ES 04704.1c 8.1.2)
    /// - 100% 시료에서 0~10% 영향 → TU = 0
    /// - 100% 시료에서 10~49% 영향 → TU = 0.02 × 유영저해율(%)
    /// - 100% 시료에서 51~99% 영향 → 추가 희석 필요 (TU 산출 불가 표시)
    /// </summary>
    public static EcotoxResult CalculateFallbackTU(int organismsAt100, int mortalitiesAt100)
    {
        double rate = organismsAt100 > 0 ? (double)mortalitiesAt100 / organismsAt100 * 100 : 0;

        if (rate <= 10)
            return new EcotoxResult(0, 0, 0, 0, "직접산출", 0, false,
                $"100% 시료 유영저해율 {rate:F1}% (10% 이하) → TU = 0");

        if (rate < 50)
        {
            double tu = Math.Round(0.02 * rate, 1);
            return new EcotoxResult(0, 0, 0, tu, "직접산출", 0, false,
                $"100% 시료 유영저해율 {rate:F1}% → TU = 0.02 × {rate:F1} = {tu}");
        }

        // 51~99%: 추가 희석 필요
        return new EcotoxResult(0, 0, 0, 0, "산출불가", 0, false,
            $"100% 시료 유영저해율 {rate:F1}% (51~99%) → 100%~50% 사이 추가 희석 필요");
    }

    /// <summary>대조군 유효성 검증 (사망률 10% 초과 시 시험 무효)</summary>
    public static string? ValidateControl(int controlOrganisms, int controlMortalities)
    {
        if (controlOrganisms <= 0) return "대조군 생물수가 0입니다.";
        double rate = (double)controlMortalities / controlOrganisms;
        if (rate > MaxControlMortality)
            return $"대조군 사망률 {rate * 100:F1}%로 10% 초과 — 시험 무효 (ES 04704.1c)";
        return null;
    }

    /// <summary>Probit 적용 가능 여부 확인 (1~99% 반응 데이터 2개 이상)</summary>
    public static bool CanUseProbit(int[] organisms, int[] mortalities)
    {
        int partialCount = 0;
        for (int i = 0; i < organisms.Length; i++)
        {
            if (organisms[i] <= 0) continue;
            double p = (double)mortalities[i] / organisms[i];
            if (p > 0.01 && p < 0.99) partialCount++;
        }
        return partialCount >= 2;
    }

    /// <summary>TSK 적용 가능 여부 확인 (유영저해/사망률 자료 1개 이상)</summary>
    public static bool CanUseTSK(int[] organisms, int[] mortalities)
    {
        for (int i = 0; i < organisms.Length; i++)
            if (organisms[i] > 0 && mortalities[i] > 0) return true;
        return false;
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  TSK (Trimmed Spearman-Karber)
    // ══════════════════════════════════════════════════════════════════════════

    /// <param name="concentrations">농도(%) 배열 (대조군 제외, 오름차순)</param>
    /// <param name="organisms">각 농도별 생물수 (반복구 합계, 보통 20)</param>
    /// <param name="mortalities">각 농도별 유영저해+치사 합계</param>
    /// <param name="controlOrganisms">대조군 생물수</param>
    /// <param name="controlMortalities">대조군 유영저해+치사수</param>
    /// <param name="trimPercent">trim %, null이면 자동</param>
    public static EcotoxResult CalculateTSK(
        double[] concentrations, int[] organisms, int[] mortalities,
        int controlOrganisms, int controlMortalities,
        double? trimPercent = null)
    {
        int n = concentrations.Length;
        if (n < 1) throw new ArgumentException("최소 1개 농도가 필요합니다 (TSK).");
        if (organisms.Length != n || mortalities.Length != n)
            throw new ArgumentException("농도, 생물수, 사망수 배열 길이가 일치해야 합니다.");

        // 대조군 검증
        var ctrlWarn = ValidateControl(controlOrganisms, controlMortalities);

        // TSK 적용 가능 여부
        if (!CanUseTSK(organisms, mortalities))
            return CalculateFallbackTU(
                organisms.Length > 0 ? organisms[^1] : 0,
                mortalities.Length > 0 ? mortalities[^1] : 0);

        // 1. 반응 비율 계산
        var props = new double[n];
        for (int i = 0; i < n; i++)
            props[i] = organisms[i] > 0 ? (double)mortalities[i] / organisms[i] : 0;

        // 2. Abbott 보정 (대조군 사망률)
        double pc = controlOrganisms > 0 ? (double)controlMortalities / controlOrganisms : 0;
        if (pc > 0)
        {
            for (int i = 0; i < n; i++)
            {
                props[i] = (props[i] - pc) / (1 - pc);
                props[i] = Math.Clamp(props[i], 0, 1);
            }
        }

        // 3. 단조증가 보정 (isotonic regression — PAVA)
        bool smoothed = false;
        var sm = (double[])props.Clone();
        for (int i = 1; i < n; i++)
        {
            if (sm[i] < sm[i - 1])
            {
                smoothed = true;
                double sum = sm[i - 1] + sm[i];
                int cnt = 2;
                int j = i - 1;
                while (j > 0 && sum / cnt < sm[j - 1])
                {
                    sum += sm[j - 1]; cnt++; j--;
                }
                double avg = sum / cnt;
                for (int k = j; k <= i; k++) sm[k] = avg;
            }
        }
        props = sm;

        // 4. Trim 결정
        double trim;
        if (trimPercent.HasValue)
            trim = trimPercent.Value / 100.0;
        else
        {
            double lowerTrim = props[0];
            double upperTrim = 1.0 - props[n - 1];
            trim = Math.Max(lowerTrim, upperTrim);
            trim = Math.Max(trim, 0);
        }

        // 5. log 농도
        var logConc = new double[n];
        for (int i = 0; i < n; i++)
            logConc[i] = Math.Log10(Math.Max(concentrations[i], 1e-10));

        // 6. LC50/EC50 산출 — EPA TSK 표준 방식.
        //    PAVA 평활값(props) 을 그대로 사용해 p=0.5 의 선형보간(log-농도 공간)으로 LC50 추정.
        //    이전 구현은 trim 보정을 [trim, 1-trim] → [0, 1] 로 선형 재스케일한 뒤 SK 사다리꼴 적분을
        //    적용했는데, 표준 EPA TSK.exe (Hamilton 1977) 와 결과가 달라 EC50 이 과소추정됨.
        //    trim 은 외삽 / 신뢰구간용으로만 사용하고 LC50 자체는 PAVA 평활값으로 직접 보간.
        string? warning = ctrlWarn;

        double logEC50;
        int idx = -1;
        for (int i = 0; i < n; i++)
        {
            if (props[i] >= 0.5) { idx = i; break; }
        }

        if (idx == -1)
        {
            // 모든 농도에서 p < 0.5 — LC50 가 시험농도 최고치 초과 (외삽)
            logEC50 = logConc[n - 1];
            warning = (warning ?? "") + " LC50 가 시험농도 최고치 초과 (외삽).";
        }
        else if (idx == 0)
        {
            // 최저농도에서 이미 p ≥ 0.5 — LC50 가 시험농도 최저치 미만 (외삽)
            logEC50 = logConc[0];
            warning = (warning ?? "") + " LC50 가 시험농도 최저치 미만 (외삽).";
        }
        else
        {
            double pLow  = props[idx - 1];
            double pHigh = props[idx];
            double xLow  = logConc[idx - 1];
            double xHigh = logConc[idx];
            if (pHigh - pLow < 1e-12)
                logEC50 = xHigh;
            else
                logEC50 = xLow + (0.5 - pLow) / (pHigh - pLow) * (xHigh - xLow);
        }

        // 7. 분산 계산 — PAVA 평활값(props) 기준
        double variance = 0;
        for (int i = 0; i < n; i++)
        {
            double dx = i < n - 1 ? logConc[i + 1] - logConc[i]
                       : i > 0   ? logConc[i] - logConc[i - 1] : 0.3;
            if (organisms[i] > 1)
                variance += dx * dx * props[i] * (1 - props[i]) / (organisms[i] - 1);
        }
        double se = Math.Sqrt(variance);

        // 8. 결과
        double ec50 = Math.Pow(10, logEC50);
        double lower = Math.Pow(10, logEC50 - 1.96 * se);
        double upper = Math.Pow(10, logEC50 + 1.96 * se);
        double tu = ec50 > 0 ? Math.Round(100.0 / ec50, 1) : 0;

        return new EcotoxResult(
            Math.Round(ec50, 2), Math.Round(lower, 2), Math.Round(upper, 2),
            tu, "TSK", Math.Round(trim * 100, 2), smoothed, warning);
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  그래프법 (ES 04704.1c 7.6조)
    //  상용로그-유영저해율 관계 그래프에서 50% 반응점 선형보간
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 그래프법 EC50 산출 (ES 04704.1c 7.6조 후단)
    /// - 시험농도 상용로그 vs 유영저해율 관계에서 50% 교점을 선형보간으로 산출
    /// - 95% 신뢰구간은 제공하지 않음
    /// - TSK/Probit 적용 불가 또는 담당자 판단에 따라 사용
    /// </summary>
    public static EcotoxResult CalculateGraphical(
        double[] concentrations, int[] organisms, int[] mortalities,
        int controlOrganisms, int controlMortalities)
    {
        int n = concentrations.Length;
        if (n < 2) throw new ArgumentException("최소 2개 농도가 필요합니다 (그래프법).");
        if (organisms.Length != n || mortalities.Length != n)
            throw new ArgumentException("농도, 생물수, 사망수 배열 길이가 일치해야 합니다.");

        var ctrlWarn = ValidateControl(controlOrganisms, controlMortalities);

        // 반응 비율 + Abbott 보정
        double pc = controlOrganisms > 0 ? (double)controlMortalities / controlOrganisms : 0;
        var props = new double[n];
        for (int i = 0; i < n; i++)
        {
            double p = organisms[i] > 0 ? (double)mortalities[i] / organisms[i] : 0;
            if (pc > 0) p = (p - pc) / (1 - pc);
            props[i] = Math.Clamp(p, 0, 1);
        }

        // 최저농도에서 이미 50% 초과 → ES 8.1.2.4 (TU > 100/최저농도)
        if (props[0] >= 0.5)
        {
            double tuMin = concentrations[0] > 0 ? 100.0 / concentrations[0] : 0;
            return new EcotoxResult(0, 0, 0, 0, "그래프법", 0, false,
                $"최저농도 {concentrations[0]}%에서 이미 {props[0] * 100:F1}% 영향 → EC50 < {concentrations[0]}% (TU > {tuMin:F1})");
        }

        // 50% 반응이 걸치는 구간 찾기 (낮은 농도 → 높은 농도 순)
        int idx = -1;
        for (int i = 0; i < n - 1; i++)
        {
            if (props[i] < 0.5 && props[i + 1] >= 0.5) { idx = i; break; }
        }

        // 최고농도에서도 50% 미달 → ES 8.1.2 직접산출 규칙으로 폴백
        if (idx < 0)
            return CalculateFallbackTU(organisms[n - 1], mortalities[n - 1]);

        // 상용로그 - 반응율 선형보간
        double logCLo = Math.Log10(Math.Max(concentrations[idx], 1e-10));
        double logCHi = Math.Log10(Math.Max(concentrations[idx + 1], 1e-10));
        double pLo = props[idx];
        double pHi = props[idx + 1];
        double logEC50 = logCLo + (0.5 - pLo) / (pHi - pLo) * (logCHi - logCLo);
        double ec50 = Math.Pow(10, logEC50);
        double tu = ec50 > 0 ? Math.Round(100.0 / ec50, 1) : 0;

        string msg = $"상용로그-유영저해율 그래프법 선형보간: "
                   + $"{concentrations[idx]}%({pLo * 100:F1}%) ↔ {concentrations[idx + 1]}%({pHi * 100:F1}%)"
                   + " | 95% CI 산출 불가";
        if (!string.IsNullOrEmpty(ctrlWarn)) msg += $"\n{ctrlWarn}";

        return new EcotoxResult(
            Math.Round(ec50, 2), 0, 0,
            tu, "그래프법", 0, false, msg);
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  Probit 분석
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Probit 분석 (가중최소제곱 회귀)
    /// ES 04704.1c: 1~99% 사이 유영저해/사망 데이터 2개 이상 필요
    /// </summary>
    public static EcotoxResult CalculateProbit(
        double[] concentrations, int[] organisms, int[] mortalities,
        int controlOrganisms, int controlMortalities)
    {
        int n = concentrations.Length;
        if (n < 2) throw new ArgumentException("최소 2개 농도가 필요합니다.");

        var ctrlWarn = ValidateControl(controlOrganisms, controlMortalities);

        if (!CanUseProbit(organisms, mortalities))
            throw new InvalidOperationException(
                "Probit 적용 불가: 1~99% 반응 데이터가 2개 미만 (TSK 사용 권장)");

        // 1. 반응 비율 + Abbott 보정
        var props = new double[n];
        double pc = controlOrganisms > 0 ? (double)controlMortalities / controlOrganisms : 0;
        for (int i = 0; i < n; i++)
        {
            double p = organisms[i] > 0 ? (double)mortalities[i] / organisms[i] : 0;
            if (pc > 0) p = (p - pc) / (1 - pc);
            if (p <= 0) p = 0.25 / organisms[i];
            if (p >= 1) p = (organisms[i] - 0.25) / organisms[i];
            props[i] = Math.Clamp(p, 0.001, 0.999);
        }

        // 2. log 농도
        var logX = new double[n];
        for (int i = 0; i < n; i++)
            logX[i] = Math.Log10(Math.Max(concentrations[i], 1e-10));

        // 3. 반복 가중최소제곱 (IWLS)
        double a = 0, b = 1;
        for (int iter = 0; iter < 25; iter++)
        {
            double sumW = 0, sumWX = 0, sumWY = 0, sumWXX = 0, sumWXY = 0;
            for (int i = 0; i < n; i++)
            {
                double probit = NormInv(props[i]) + 5.0;
                double w = organisms[i] * props[i] * (1 - props[i]);
                if (w < 0.001) w = 0.001;

                sumW   += w;
                sumWX  += w * logX[i];
                sumWY  += w * probit;
                sumWXX += w * logX[i] * logX[i];
                sumWXY += w * logX[i] * probit;
            }

            double denom = sumW * sumWXX - sumWX * sumWX;
            if (Math.Abs(denom) < 1e-15) break;

            double newA = (sumWY * sumWXX - sumWXY * sumWX) / denom;
            double newB = (sumW * sumWXY - sumWX * sumWY) / denom;

            if (Math.Abs(newA - a) < 1e-8 && Math.Abs(newB - b) < 1e-8) break;
            a = newA; b = newB;

            for (int i = 0; i < n; i++)
            {
                double pExpected = NormCDF(a + b * logX[i] - 5.0);
                double rawP = organisms[i] > 0 ? (double)mortalities[i] / organisms[i] : 0;
                if (pc > 0) rawP = (rawP - pc) / (1 - pc);
                props[i] = Math.Clamp(0.5 * rawP + 0.5 * pExpected, 0.001, 0.999);
            }
        }

        // 4. EC50 계산
        if (Math.Abs(b) < 1e-10) b = 0.001;
        double logEC50 = (5.0 - a) / b;
        double ec50 = Math.Pow(10, logEC50);

        // 5. 신뢰구간 (분산-공분산 행렬)
        double sW = 0, sWX = 0, sWXX = 0;
        for (int i = 0; i < n; i++)
        {
            double w = organisms[i] * props[i] * (1 - props[i]);
            if (w < 0.001) w = 0.001;
            sW += w; sWX += w * logX[i]; sWXX += w * logX[i] * logX[i];
        }
        double det = sW * sWXX - sWX * sWX;
        double varA = det > 1e-15 ? sWXX / det : 1;
        double varB = det > 1e-15 ? sW / det : 1;
        double covAB = det > 1e-15 ? -sWX / det : 0;

        double seLog = Math.Sqrt(
            varA / (b * b) +
            ((5.0 - a) * (5.0 - a)) / (b * b * b * b) * varB -
            2.0 * (5.0 - a) / (b * b * b) * covAB);

        double lower = Math.Pow(10, logEC50 - 1.96 * seLog);
        double upper = Math.Pow(10, logEC50 + 1.96 * seLog);
        double tu = ec50 > 0 ? Math.Round(100.0 / ec50, 1) : 0;

        return new EcotoxResult(
            Math.Round(ec50, 2), Math.Round(lower, 2), Math.Round(upper, 2),
            tu, "Probit", 0, false, ctrlWarn);
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  통계 유틸리티
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>표준정규분포 CDF</summary>
    private static double NormCDF(double x)
    {
        const double a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741;
        const double a4 = -1.453152027, a5 = 1.061405429, p = 0.3275911;
        double sign = x < 0 ? -1 : 1;
        x = Math.Abs(x) / Math.Sqrt(2);
        double t = 1.0 / (1.0 + p * x);
        double y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.Exp(-x * x);
        return 0.5 * (1.0 + sign * y);
    }

    /// <summary>표준정규분포 역함수 (Peter Acklam 근사)</summary>
    private static double NormInv(double p)
    {
        if (p <= 0) return -8;
        if (p >= 1) return 8;

        const double a1 = -3.969683028665376e+01, a2 = 2.209460984245205e+02;
        const double a3 = -2.759285104469687e+02, a4 = 1.383577518672690e+02;
        const double a5 = -3.066479806614716e+01, a6 = 2.506628277459239e+00;
        const double b1 = -5.447609879822406e+01, b2 = 1.615858368580409e+02;
        const double b3 = -1.556989798598866e+02, b4 = 6.680131188771972e+01;
        const double b5 = -1.328068155288572e+01;
        const double c1 = -7.784894002430293e-03, c2 = -3.223964580411365e-01;
        const double c3 = -2.400758277161838e+00, c4 = -2.549732539343734e+00;
        const double c5 = 4.374664141464968e+00, c6 = 2.938163982698783e+00;
        const double d1 = 7.784695709041462e-03, d2 = 3.224671290700398e-01;
        const double d3 = 2.445134137142996e+00, d4 = 3.754408661907416e+00;
        const double pLow = 0.02425, pHigh = 1 - pLow;

        double q, r;
        if (p < pLow)
        {
            q = Math.Sqrt(-2 * Math.Log(p));
            return (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) /
                   ((((d1 * q + d2) * q + d3) * q + d4) * q + 1);
        }
        if (p <= pHigh)
        {
            q = p - 0.5; r = q * q;
            return (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q /
                   (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1);
        }
        q = Math.Sqrt(-2 * Math.Log(1 - p));
        return -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) /
               ((((d1 * q + d2) * q + d3) * q + d4) * q + 1);
    }
}
