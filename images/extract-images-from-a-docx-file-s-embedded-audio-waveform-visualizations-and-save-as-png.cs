using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class ExtractWaveformImages
{
    public static void Main()
    {
        // Define file names.
        const string waveformImagePath = "waveform.png";
        const string documentPath = "sample.docx";

        // -------------------------------------------------
        // Step 1: Create a sample waveform image.
        // -------------------------------------------------
        const int imgWidth = 300;
        const int imgHeight = 100;
        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);

        // Draw a simple sine‑wave like pattern to simulate an audio waveform.
        using (Pen pen = new Pen(Color.Black, 2))
        {
            Point previousPoint = new Point(0, imgHeight / 2);
            for (int x = 1; x < imgWidth; x++)
            {
                double radians = 2 * Math.PI * x / imgWidth;
                int y = (int)(imgHeight / 2 + (imgHeight / 4) * Math.Sin(radians));
                Point currentPoint = new Point(x, y);
                graphics.DrawLine(pen, previousPoint, currentPoint);
                previousPoint = currentPoint;
            }
        }

        // Save the generated image to the local file system.
        bitmap.Save(waveformImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // Step 2: Create a DOCX document and embed the image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(waveformImagePath);
        doc.Save(documentPath);

        // -------------------------------------------------
        // Step 3: Load the document and extract all images.
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outputFileName = $"extracted_{extractedCount}{extension}";

                // Save the image to the file system.
                shape.ImageData.Save(outputFileName);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: indicate success.
        Console.WriteLine($"Successfully extracted {extractedCount} image(s).");
    }
}
