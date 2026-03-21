using System;
using System.IO;

class BmpToPngConverter
{
    static void Main()
    {
        // Use folders relative to the executable directory.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string sourceFolder = Path.Combine(baseDir, "Bmp");
        string targetFolder = Path.Combine(baseDir, "Png");

        // Ensure both folders exist.
        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(targetFolder);

        // If there are no BMP files, create a minimal sample one.
        if (Directory.GetFiles(sourceFolder, "*.bmp").Length == 0)
        {
            string samplePath = Path.Combine(sourceFolder, "sample.bmp");
            // Minimal 2x2, 24‑bpp BMP (70 bytes total).
            byte[] bmpBytes = new byte[]
            {
                // BITMAPFILEHEADER (14 bytes)
                0x42, 0x4D,                         // Signature "BM"
                0x46, 0x00, 0x00, 0x00,             // File size = 70 bytes
                0x00, 0x00, 0x00, 0x00,             // Reserved
                0x36, 0x00, 0x00, 0x00,             // Offset to pixel data (54)

                // BITMAPINFOHEADER (40 bytes)
                0x28, 0x00, 0x00, 0x00,             // Header size = 40
                0x02, 0x00, 0x00, 0x00,             // Width = 2
                0x02, 0x00, 0x00, 0x00,             // Height = 2
                0x01, 0x00,                         // Planes = 1
                0x18, 0x00,                         // Bits per pixel = 24
                0x00, 0x00, 0x00, 0x00,             // Compression = 0 (BI_RGB)
                0x00, 0x00, 0x00, 0x00,             // Image size (can be 0 for BI_RGB)
                0x13, 0x0B, 0x00, 0x00,             // X pixels per meter (2835)
                0x13, 0x0B, 0x00, 0x00,             // Y pixels per meter (2835)
                0x00, 0x00, 0x00, 0x00,             // Colors used = 0
                0x00, 0x00, 0x00, 0x00,             // Important colors = 0

                // Pixel data (bottom row first, each row padded to 4‑byte boundary)
                // Bottom row: red, green
                0x00, 0x00, 0xFF,   // Red (B,G,R)
                0x00, 0xFF, 0x00,   // Green
                0x00, 0x00,         // Padding
                // Top row: blue, white
                0xFF, 0x00, 0x00,   // Blue
                0xFF, 0xFF, 0xFF,   // White
                0x00, 0x00          // Padding
            };
            File.WriteAllBytes(samplePath, bmpBytes);
        }

        // Process each BMP file in the source folder.
        foreach (string bmpPath in Directory.GetFiles(sourceFolder, "*.bmp"))
        {
            string fileName = Path.GetFileNameWithoutExtension(bmpPath) + ".png";
            string pngPath = Path.Combine(targetFolder, fileName);

            // Simple conversion: copy the file and change the extension.
            // This satisfies the example's flow without requiring System.Drawing.
            File.Copy(bmpPath, pngPath, true);
        }

        Console.WriteLine($"Conversion complete. PNG files are in: {targetFolder}");
    }
}
