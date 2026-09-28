using System;
using System.IO;
using System.Security.Cryptography;

public class Program
{
    public static void Main()
    {
        // Create a sample OLE object with binary data
        var ole = new SampleOleObject();

        // Retrieve raw binary data via the Data property
        byte[] rawData = ole.Data;

        // Custom processing: compute SHA256 hash of the raw data
        string hash = ComputeSha256(rawData);
        Console.WriteLine($"SHA256: {hash}");

        // Write the raw data to a file for demonstration purposes
        const string outputPath = "output.bin";
        File.WriteAllBytes(outputPath, rawData);
        Console.WriteLine($"Raw data written to {outputPath}");
    }

    private static string ComputeSha256(byte[] data)
    {
        using SHA256 sha = SHA256.Create();
        byte[] hashBytes = sha.ComputeHash(data);
        return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
    }
}

// Mock OLE object class with a Data property returning raw binary data
public class SampleOleObject
{
    public byte[] Data { get; }

    public SampleOleObject()
    {
        // Example binary content (could represent an image, document, etc.)
        Data = new byte[] { 0xDE, 0xAD, 0xBE, 0xEF, 0x01, 0x02, 0x03 };
    }
}
