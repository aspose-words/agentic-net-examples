using System;
using System.IO;

public class OleObjectSaver
{
    public static void Main()
    {
        // Sample raw binary data representing an OLE object (OLE Compound File header)
        byte[] oleData = new byte[]
        {
            0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1,
            // Additional dummy bytes for illustration
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
        };

        // Create a temporary file path
        string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".bin");

        // Write the raw OLE data to the temporary file
        File.WriteAllBytes(tempFilePath, oleData);

        // Output the location of the saved file (optional)
        Console.WriteLine($"OLE data saved to: {tempFilePath}");
    }
}
