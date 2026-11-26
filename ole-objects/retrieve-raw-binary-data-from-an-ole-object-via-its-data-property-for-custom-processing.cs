using System;
using System.Security.Cryptography;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Simulated OLE object containing raw binary data
        byte[] sampleData = Encoding.UTF8.GetBytes("Example binary content for OLE object.");
        OLEObject ole = new OLEObject(sampleData);

        // Retrieve raw binary data via the Data property
        byte[] rawData = ole.Data;

        // Custom processing: compute SHA-256 hash of the raw data
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hash = sha256.ComputeHash(rawData);
            string hashHex = BitConverter.ToString(hash).Replace("-", "");
            Console.WriteLine($"SHA-256 hash of OLE data: {hashHex}");
        }

        // Additional custom processing: output the raw bytes as a hex string
        StringBuilder sb = new StringBuilder();
        foreach (byte b in rawData)
        {
            sb.Append(b.ToString("X2"));
        }
        Console.WriteLine($"Raw OLE data (hex): {sb}");
    }
}

// Simple representation of an OLE object with a Data property
public class OLEObject
{
    public byte[] Data { get; }

    public OLEObject(byte[] data)
    {
        Data = data ?? Array.Empty<byte>();
    }
}
