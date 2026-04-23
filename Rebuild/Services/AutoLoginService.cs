using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace ETA.Rebuild.Services;

public record AutoLoginPayload(string Id, string Pw, int Version);

public static class AutoLoginService
{
    private static string FilePath
    {
        get
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ETA");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "autologin.dat");
        }
    }

    public static void Save(string id, string pw, int version)
    {
        try
        {
            var payload = new AutoLoginPayload(id, pw, version);
            var json = JsonSerializer.Serialize(payload);
            var cipher = Encrypt(json);
            File.WriteAllText(FilePath, cipher);
        }
        catch { }
    }

    public static AutoLoginPayload? Load()
    {
        try
        {
            if (!File.Exists(FilePath)) return null;
            var cipher = File.ReadAllText(FilePath);
            if (string.IsNullOrWhiteSpace(cipher)) return null;
            var json = Decrypt(cipher);
            return JsonSerializer.Deserialize<AutoLoginPayload>(json);
        }
        catch { return null; }
    }

    public static void Clear()
    {
        try { if (File.Exists(FilePath)) File.Delete(FilePath); } catch { }
    }

    private static byte[] DeriveKey()
    {
        var seed = Environment.UserName + "|" + Environment.MachineName + "|ETA-AUTOLOGIN-v1";
        using var sha = SHA256.Create();
        return sha.ComputeHash(Encoding.UTF8.GetBytes(seed));
    }

    private static string Encrypt(string plaintext)
    {
        var key = DeriveKey();
        var nonce = RandomNumberGenerator.GetBytes(12);
        var data = Encoding.UTF8.GetBytes(plaintext);
        var cipher = new byte[data.Length];
        var tag = new byte[16];
        using var aes = new AesGcm(key, 16);
        aes.Encrypt(nonce, data, cipher, tag);
        var combined = new byte[nonce.Length + tag.Length + cipher.Length];
        Buffer.BlockCopy(nonce,  0, combined, 0,                        nonce.Length);
        Buffer.BlockCopy(tag,    0, combined, nonce.Length,             tag.Length);
        Buffer.BlockCopy(cipher, 0, combined, nonce.Length + tag.Length, cipher.Length);
        return Convert.ToBase64String(combined);
    }

    private static string Decrypt(string base64)
    {
        var combined = Convert.FromBase64String(base64);
        var nonce  = new byte[12];
        var tag    = new byte[16];
        var cipher = new byte[combined.Length - 28];
        Buffer.BlockCopy(combined, 0,  nonce,  0, 12);
        Buffer.BlockCopy(combined, 12, tag,    0, 16);
        Buffer.BlockCopy(combined, 28, cipher, 0, cipher.Length);
        var plain = new byte[cipher.Length];
        using var aes = new AesGcm(DeriveKey(), 16);
        aes.Decrypt(nonce, cipher, tag, plain);
        return Encoding.UTF8.GetString(plain);
    }
}
