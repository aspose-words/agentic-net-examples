using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class ExtractVideoFrameImages
{
    public static void Main()
    {
        // Create a high‑resolution sample image that will simulate a video frame.
        const int width = 1920;
        const int height = 1080;
        const string sampleImagePath = "frame.png";

        // Generate the sample bitmap.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        // Draw a simple rectangle so the image is recognizable.
        using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 10))
        {
            graphics.DrawRectangle(pen, 100, 100, width - 200, height - 200);
        }

        // Save the bitmap to a file that will be inserted into the document.
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // Create a new DOCX document and insert the sample image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Reload the document (simulating an existing file with embedded video frames).
        Document loadedDoc = new Document(docPath);

        // Extract all images (including possible video‑frame thumbnails) and save them as PNG files.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            string outputPath = $"extracted-{extractedCount + 1}.png";

            // Save the image data directly to a PNG file.
            shape.ImageData.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to save extracted image to '{outputPath}'.");

            extractedCount++;
        }

        // Ensure that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional cleanup (commented out; uncomment if you want to delete temporary files).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
