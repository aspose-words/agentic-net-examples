using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image that will act as a map.
        const string mapImagePath = "map.png";
        const int mapWidth = 800;
        const int mapHeight = 600;

        using (Bitmap bitmap = new Bitmap(mapWidth, mapHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);
                // Draw a simple rectangle to simulate map content.
                graphics.DrawRectangle(Pens.Black, 50, 50, mapWidth - 100, mapHeight - 100);
            }
            // Save the image as PNG.
            bitmap.Save(mapImagePath, ImageFormat.Png);
        }

        // Create a DOCX document and insert the sample map image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(mapImagePath);
        doc.Save(docPath);

        // Load the document and extract all embedded images (including map objects).
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outputImagePath = $"extracted-{extractedCount}.png";
                shape.ImageData.Save(outputImagePath);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
        {
            throw new Exception("No images were extracted from the document.");
        }

        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved as PNG files.");
    }
}
