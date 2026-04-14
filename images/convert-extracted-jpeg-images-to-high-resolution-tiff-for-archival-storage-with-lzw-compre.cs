using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image using Aspose.Drawing
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.White);
            // Draw a simple rectangle
            g.DrawRectangle(new Pen(Color.Blue, 5), 20, 20, 160, 160);
            // Save as JPEG
            bitmap.Save(jpegPath);
        }

        // 2. Create a Word document and insert the JPEG image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape insertedShape = builder.InsertImage(jpegPath);
        // Ensure the shape was added and contains an image
        if (!insertedShape.HasImage)
            throw new Exception("Failed to insert image into the document.");

        // 3. Extract the image data from the shape
        ImageData imageData = insertedShape.ImageData;
        using (MemoryStream imageStream = new MemoryStream())
        {
            imageData.Save(imageStream);
            imageStream.Position = 0; // Reset before reuse

            // 4. Create a new document that contains only the extracted image
            Document imageDoc = new Document();
            DocumentBuilder imgBuilder = new DocumentBuilder(imageDoc);
            // Insert image from the extracted stream
            Shape imgShape = imgBuilder.InsertImage(imageStream.ToArray());
            if (!imgShape.HasImage)
                throw new Exception("Failed to insert extracted image into the new document.");

            // 5. Save the document as a high‑resolution TIFF with LZW compression
            string tiffPath = Path.Combine(artifactsDir, "converted.tiff");
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Lzw,
                Resolution = 300 // High resolution (300 DPI)
            };
            imageDoc.Save(tiffPath, tiffOptions);

            // 6. Validate that the TIFF file was created
            if (!File.Exists(tiffPath))
                throw new Exception("TIFF conversion failed; output file not found.");
        }
    }
}
