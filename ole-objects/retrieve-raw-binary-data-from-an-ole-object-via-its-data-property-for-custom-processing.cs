using System;
using System.Security.Cryptography;

public class Program
{
    public static void Main()
    {
        // Simulated OLE object containing raw binary data (e.g., "Hello World")
        var oleObject = new OleObjectMock(new byte[] { 0x48, 0x65, 0x6C, 0x6C, 0x6F, 0x20, 0x57, 0x6F, 0x72, 0x6C, 0x64 });

        // Retrieve raw binary data via the Data property
        byte[] rawData = oleObject.Data;

        // Custom processing: compute SHA256 hash of the binary data
        string hash = ComputeSha256Hash(rawData);

        Console.WriteLine($"SHA256 hash of OLE data: {hash}");
    }

    private static string ComputeSha256Hash(byte[] data)
    {
        using var sha256 = SHA256.Create();
        byte[] hashBytes = sha256.ComputeHash(data);
        return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
    }
}

public class OleObjectMock
{
    private readonly byte[] _data;

    public OleObjectMock(byte[] data)
    {
        _data = data;
    }

    // Exposes the raw binary data of the OLE object
    public byte[] Data => _data;
}
