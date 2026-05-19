using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create sample GIF image.
        string gifPath = Path.Combine(artifactsDir, "sample.gif");
        using (Bitmap gifBitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(gifBitmap))
            {
                g.Clear(Color.LightBlue);
            }
            gifBitmap.Save(gifPath, ImageFormat.Gif);
        }

        // Create sample PNG image (replacement).
        string pngPath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap pngBitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(pngBitmap))
            {
                g.Clear(Color.LightGreen);
            }
            pngBitmap.Save(pngPath, ImageFormat.Png);
        }

        // Verify that the images were created.
        if (!File.Exists(gifPath) || !File.Exists(pngPath))
            throw new Exception("Failed to create sample images.");

        // Create a Word document containing GIF images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with GIF images:");
        builder.InsertImage(gifPath);
        builder.InsertParagraph();
        builder.InsertImage(gifPath);
        string originalDocPath = Path.Combine(artifactsDir, "Original.docx");
        doc.Save(originalDocPath);

        // Load the document for processing.
        Document loadedDoc = new Document(originalDocPath);

        // Replace every GIF image with the corresponding PNG image.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                // Use the custom mapping (GIF -> PNG).
                shape.ImageData.SetImage(pngPath);
            }
        }

        // Save the modified document.
        string modifiedDocPath = Path.Combine(artifactsDir, "Modified.docx");
        loadedDoc.Save(modifiedDocPath);

        // Validate that the output file exists.
        if (!File.Exists(modifiedDocPath))
            throw new Exception("Modified document was not saved.");

        // Ensure that at least one PNG image is present after replacement.
        int pngCount = 0;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
                pngCount++;
        }

        if (pngCount == 0)
            throw new Exception("No PNG images were found after replacement.");

        // Optional: indicate success.
        Console.WriteLine("GIF images successfully replaced with PNG equivalents.");
    }
}
