using System;
using System.IO;

public class Program
{
    public static void Main(string[] args)
    {
        // Simulated raw binary data of an OLE object (OLE Compound File header)
        byte[] oleData = new byte[]
        {
            0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            // Additional dummy bytes for illustration
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08
        };

        // Create a temporary file path
        string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".bin");

        try
        {
            // Write the raw OLE data to the temporary file
            File.WriteAllBytes(tempFilePath, oleData);

            // Output the location of the temporary file (for external analysis)
            Console.WriteLine($"OLE data saved to temporary file: {tempFilePath}");
        }
        finally
        {
            // Clean up: delete the temporary file if it exists
            if (File.Exists(tempFilePath))
            {
                try
                {
                    File.Delete(tempFilePath);
                }
                catch
                {
                    // If deletion fails, ignore to avoid crashing the program
                }
            }
        }
    }
}
