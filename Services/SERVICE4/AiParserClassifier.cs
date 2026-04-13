using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ETA.Services.SERVICE4;

/// <summary>
/// ONNX 기반 분석항목 파서 자동 분류기
/// Data/분석항목_분류.onnx 로드 → 파일 텍스트 → 파서 키 반환
/// </summary>
public static class AiParserClassifier
{
    private static InferenceSession? _session;
    private static readonly string OnnxPath = Path.Combine(AppContext.BaseDirectory, "Data", "분석항목_분류.onnx");
    private static readonly object _lock = new();

    // ── 모델 준비 여부 ──────────────────────────────────────────────────
    public static bool IsModelReady() => File.Exists(OnnxPath);

    // ── 모델 언로드 (재훈련 후 갱신용) ────────────────────────────────
    public static void Reload()
    {
        lock (_lock)
        {
            _session?.Dispose();
            _session = null;
        }
    }

    // ── 분류 예측 ──────────────────────────────────────────────────────
    /// <returns>파서 키 (예: "GC", "BOD") 또는 null (모델 없음/오류)</returns>
    public static string? Predict(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        if (!IsModelReady()) return null;

        try
        {
            var session = GetSession();
            if (session == null) return null;

            // string tensor [1, 1]
            var tensor = new DenseTensor<string>(new[] { text }, new[] { 1, 1 });
            var inputs = new List<NamedOnnxValue>
            {
                NamedOnnxValue.CreateFromTensor("string_input", tensor),
            };

            using var results = session.Run(inputs);

            // 첫 번째 출력이 label
            var labelTensor = results.First().AsTensor<string>();
            return labelTensor.First();
        }
        catch (Exception ex)
        {
            return null;
        }
    }

    /// <returns>(파서 키, 확률) 상위 3개</returns>
    public static List<(string Label, float Prob)> PredictTopK(string text, int k = 3)
    {
        var result = new List<(string, float)>();
        if (string.IsNullOrWhiteSpace(text)) return result;
        if (!IsModelReady()) return result;

        try
        {
            var session = GetSession();
            if (session == null) return result;

            var tensor = new DenseTensor<string>(new[] { text }, new[] { 1, 1 });
            var inputs = new List<NamedOnnxValue>
            {
                NamedOnnxValue.CreateFromTensor("string_input", tensor),
            };

            using var results = session.Run(inputs);

            // 두 번째 출력이 probabilities (map<string, float>)
            var probResult = results.Skip(1).FirstOrDefault();
            if (probResult?.Value is IEnumerable<IDictionary<string, float>> probMaps)
            {
                var map = probMaps.First();
                result = map
                    .OrderByDescending(kv => kv.Value)
                    .Take(k)
                    .Select(kv => (kv.Key, kv.Value))
                    .ToList();
            }
        }
        catch (Exception ex)
        {
        }

        return result;
    }

    // ── 세션 Lazy 로드 ─────────────────────────────────────────────────
    private static InferenceSession? GetSession()
    {
        if (_session != null) return _session;
        lock (_lock)
        {
            if (_session != null) return _session;
            try
            {
                _session = new InferenceSession(OnnxPath);
            }
            catch (Exception ex)
            {
            }
        }
        return _session;
    }
}
