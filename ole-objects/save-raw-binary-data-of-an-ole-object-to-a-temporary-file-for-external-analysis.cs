using System;
using System.IO;

public class Program
{
    public static void Main(string[] args)
    {
        // Sample raw binary data representing an OLE object.
        byte[] oleData = new byte[]
        {
            0xD0, 0xCF, 0x11, 0xE0, // OLE header signature
            0xA1, 0xB1, 0x1A, 0xE1,
            0x00, 0x00, 0x00, 0x00,
            // ... additional bytes could follow
        };

        // Create a unique temporary file path.
        string tempFilePath = Path.Combine(
            Path.GetTempPath(),
            $"OleObject_{Guid.NewGuid():N}.bin"
        );

        // Write the binary data to the temporary file.
        File.WriteAllBytes(tempFilePath, oleData);

        // Output the location of the saved file.
        Console.WriteLine($"OLE object saved to: {tempFilePath}");
    }
}
