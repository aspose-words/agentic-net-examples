using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic PNG image using Aspose.Drawing.
        const string inputImagePath = "input.png";
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Ensure any previous file is removed.
        if (File.Exists(inputImagePath))
            File.Delete(inputImagePath);

        // Create bitmap and fill with a solid color.
        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightGray);
        // Save the bitmap as PNG.
        bitmap.Save(inputImagePath);
        // Clean up drawing resources.
        graphics.Dispose();
        bitmap.Dispose();

        // Build a Word document that contains the sample PNG image twice.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        builder.InsertParagraph();
        builder.InsertImage(inputImagePath);

        // Save the document (optional, demonstrates load/save lifecycle).
        const string docPath = "sample.docx";
        if (File.Exists(docPath))
            File.Delete(docPath);
        doc.Save(docPath);

        // Load the document (could reuse the same instance, but follows load rule).
        Document loadedDoc = new Document(docPath);

        // Extract all PNG images, apply contrast enhancement, and save them.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Apply maximum contrast (value range 0.0 – 1.0).
            shape.ImageData.Contrast = 1.0;

            // Save the enhanced image to a deterministic file name.
            string outputFileName = $"extracted_{extractedCount}.png";
            if (File.Exists(outputFileName))
                File.Delete(outputFileName);
            shape.ImageData.Save(outputFileName);
            extractedCount++;
        }

        // Validate that at least one image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");

        // Optional: clean up the temporary document file.
        // File.Delete(docPath);
    }
}
