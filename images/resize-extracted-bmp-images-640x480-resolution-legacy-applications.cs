using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ResizeBmpImages
{
    static void Main()
    {
        // Use paths relative to the current directory so the example works everywhere.
        string inputPath = Path.Combine(Environment.CurrentDirectory, "Input.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.docx");

        // Ensure the input document exists. If not, create a simple document with a BMP image.
        if (!File.Exists(inputPath))
        {
            Document tempDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(tempDoc);

            // Create a minimal 1×1 pixel BMP image in memory.
            byte[] bmpBytes = CreateSampleBmp();

            // Insert the BMP image into the document.
            Shape shape = builder.InsertImage(bmpBytes);
            shape.Width = 100;   // arbitrary initial size
            shape.Height = 100;

            tempDoc.Save(inputPath);
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Desired pixel dimensions.
        const int targetWidthPixels = 640;
        const int targetHeightPixels = 480;

        // Default Aspose.Words DPI is 96.
        const double dpi = 96.0;

        // Convert pixel dimensions to points (1 point = 1/72 inch).
        double targetWidthPoints = targetWidthPixels * 72.0 / dpi;
        double targetHeightPoints = targetHeightPixels * 72.0 / dpi;

        // Resize all BMP images in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Bmp)
                continue;

            shape.Width = targetWidthPoints;
            shape.Height = targetHeightPoints;
        }

        // Save the modified document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Generates a minimal 1×1 pixel 24‑bit BMP image as a byte array.
    static byte[] CreateSampleBmp()
    {
        return new byte[]
        {
            0x42,0x4D,                         // Signature "BM"
            0x3E,0x00,0x00,0x00,               // File size (62 bytes)
            0x00,0x00,                         // Reserved1
            0x00,0x00,                         // Reserved2
            0x36,0x00,0x00,0x00,               // Offset to pixel data (54)
            0x28,0x00,0x00,0x00,               // DIB header size (40)
            0x01,0x00,0x00,0x00,               // Width = 1
            0x01,0x00,0x00,0x00,               // Height = 1
            0x01,0x00,                         // Planes = 1
            0x18,0x00,                         // Bits per pixel = 24
            0x00,0x00,0x00,0x00,               // Compression = 0 (none)
            0x08,0x00,0x00,0x00,               // Image size (8 bytes, padded)
            0x13,0x0B,0x00,0x00,               // X pixels per meter
            0x13,0x0B,0x00,0x00,               // Y pixels per meter
            0x00,0x00,0x00,0x00,               // Colors used
            0x00,0x00,0x00,0x00,               // Important colors
            // Pixel data (blue, green, red) = white (255,255,255)
            0xFF,0xFF,0xFF,
            // Padding to 4‑byte boundary
            0x00,0x00,0x00
        };
    }
}
