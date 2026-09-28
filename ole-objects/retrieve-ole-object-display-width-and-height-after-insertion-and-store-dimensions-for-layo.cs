using System;
using System.IO;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        // Create a temporary bitmap file to use as OLE object source
        string bmpPath = Path.Combine(Path.GetTempPath(), "sample.bmp");
        CreateSampleBitmap(bmpPath);

        // Start Word application (invisible) using late binding
        Type wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Console.WriteLine("Microsoft Word is not installed on this machine.");
            return;
        }

        dynamic wordApp = Activator.CreateInstance(wordType);
        wordApp.Visible = false;

        // Add a new document
        dynamic doc = wordApp.Documents.Add();

        // Insert the OLE object (bitmap) into the document
        dynamic range = doc.Range(0, 0);
        dynamic inlineShape = range.InlineShapes.AddOLEObject(
            "Paint.Picture",   // ClassType
            bmpPath,           // FileName
            false,             // LinkToFile
            false,             // DisplayAsIcon
            Type.Missing,      // IconFileName
            Type.Missing,      // IconIndex
            Type.Missing,      // IconLabel
            Type.Missing);     // Range

        // Retrieve display dimensions (points)
        float width = (float)inlineShape.Width;
        float height = (float)inlineShape.Height;

        // Example layout calculation: area in square points
        float area = width * height;

        Console.WriteLine($"OLE object dimensions: Width = {width} pt, Height = {height} pt, Area = {area} pt²");

        // Clean up
        doc.Close(false);
        wordApp.Quit();

        // Release COM objects
        Marshal.FinalReleaseComObject(inlineShape);
        Marshal.FinalReleaseComObject(range);
        Marshal.FinalReleaseComObject(doc);
        Marshal.FinalReleaseComObject(wordApp);

        // Delete temporary bitmap file
        if (File.Exists(bmpPath))
        {
            File.Delete(bmpPath);
        }
    }

    private static void CreateSampleBitmap(string path)
    {
        // Create a minimal 2x2 pixel 24‑bit BMP (all pixels red)
        using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
        using (var bw = new BinaryWriter(fs))
        {
            int width = 2;
            int height = 2;
            int rowSize = ((24 * width + 31) / 32) * 4; // padded to 4‑byte boundary
            int pixelArraySize = rowSize * height;
            int fileSize = 54 + pixelArraySize;

            // BITMAPFILEHEADER
            bw.Write((ushort)0x4D42);          // Signature "BM"
            bw.Write(fileSize);                // File size
            bw.Write((ushort)0);               // Reserved1
            bw.Write((ushort)0);               // Reserved2
            bw.Write(54);                      // Offset to pixel data

            // BITMAPINFOHEADER
            bw.Write(40);                      // Header size
            bw.Write(width);
            bw.Write(height);
            bw.Write((ushort)1);               // Planes
            bw.Write((ushort)24);              // Bits per pixel
            bw.Write(0);                       // Compression (none)
            bw.Write(pixelArraySize);          // Image size
            bw.Write(0);                       // X pixels per meter
            bw.Write(0);                       // Y pixels per meter
            bw.Write(0);                       // Colors used
            bw.Write(0);                       // Important colors

            // Pixel data (bottom‑up). Each pixel: B, G, R.
            // We'll make all pixels red (B=0, G=0, R=255).
            byte[] pixelRow = new byte[rowSize];
            pixelRow[0] = 0;   // B
            pixelRow[1] = 0;   // G
            pixelRow[2] = 255; // R
            // Write two rows (bottom row first)
            bw.Write(pixelRow);
            bw.Write(pixelRow);
        }
    }
}
