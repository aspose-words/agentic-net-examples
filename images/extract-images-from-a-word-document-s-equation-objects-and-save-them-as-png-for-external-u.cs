using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Math;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image using Aspose.Drawing.
        const string sampleImagePath = "sample.png";
        const int imgWidth = 100;
        const int imgHeight = 100;

        // Create bitmap and draw a simple rectangle.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        graphics.DrawRectangle(new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 2), 10, 10, imgWidth - 20, imgHeight - 20);
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // Verify that the sample image was created.
        if (!File.Exists(sampleImagePath))
            throw new FileNotFoundException("Failed to create the sample image.", sampleImagePath);

        // Create a new Word document and insert an equation followed by the sample image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple equation using a field (EQ). This creates an OfficeMath node.
        builder.InsertField("EQ \\o(\\a\\b,\\c\\d)");

        // Insert the previously created image.
        builder.InsertImage(sampleImagePath);

        // Save the document.
        doc.Save(docPath);

        // Verify that the document was saved.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Failed to save the sample document.", docPath);

        // Load the document for extraction.
        Document loadedDoc = new Document(docPath);

        // Extract images that are located within equation (OfficeMath) objects.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        int extractedCount = 0;
        int imageIndex = 0;

        foreach (OfficeMath officeMath in officeMathNodes)
        {
            NodeCollection shapeNodes = officeMath.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string outputImagePath = $"EquationImage-{imageIndex}.png";
                    shape.ImageData.Save(outputImagePath);
                    if (!File.Exists(outputImagePath))
                        throw new InvalidOperationException($"Failed to save extracted image to {outputImagePath}.");
                    extractedCount++;
                    imageIndex++;
                }
            }
        }

        // If no images were found inside equations, fallback to extracting all images in the document.
        if (extractedCount == 0)
        {
            NodeCollection allShapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in allShapes)
            {
                if (shape.HasImage)
                {
                    string outputImagePath = $"DocumentImage-{imageIndex}.png";
                    shape.ImageData.Save(outputImagePath);
                    if (!File.Exists(outputImagePath))
                        throw new InvalidOperationException($"Failed to save extracted image to {outputImagePath}.");
                    extractedCount++;
                    imageIndex++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional cleanup (commented out to keep output files).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
