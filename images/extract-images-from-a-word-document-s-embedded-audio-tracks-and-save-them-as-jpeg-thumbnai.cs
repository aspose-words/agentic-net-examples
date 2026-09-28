using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;

public class ExtractAudioTrackThumbnails
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image that will act as the audio thumbnail.
        // -----------------------------------------------------------------
        const string sampleImagePath = "input.png";
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Create a bitmap, draw a simple graphic, and save it.
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a red ellipse as placeholder content.
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    graphics.DrawEllipse(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a Word document and insert the image (simulating an audio track thumbnail).
        // -----------------------------------------------------------------
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image into the document.
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract all images (thumbnails) to JPEG files.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                // Ensure the image data is present.
                ImageData imgData = shape.ImageData;
                if (imgData != null && imgData.ImageBytes != null && imgData.ImageBytes.Length > 0)
                {
                    string outFile = $"thumbnail-{extractedCount + 1}.jpg";
                    // Save the image as JPEG.
                    imgData.Save(outFile);
                    Console.WriteLine($"Extracted thumbnail saved to: {outFile}");
                    extractedCount++;
                }
            }
        }

        // -----------------------------------------------------------------
        // 4. Validation.
        // -----------------------------------------------------------------
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No image thumbnails were extracted from the document.");
        }
        else
        {
            Console.WriteLine($"Total thumbnails extracted: {extractedCount}");
        }
    }
}
