using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create sample images (GIF and PNG) using Aspose.Drawing.
        string gifPath = Path.Combine(artifactsDir, "sample.gif");
        string pngPath = Path.Combine(artifactsDir, "sample.png");

        // Create a 100x100 bitmap with a solid color.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.LightBlue);
        // Save as GIF.
        bitmap.Save(gifPath);
        // Save as PNG.
        bitmap.Save(pngPath);
        // Clean up drawing resources.
        graphics.Dispose();
        bitmap.Dispose();

        // Verify that the image files were created.
        if (!File.Exists(gifPath) || !File.Exists(pngPath))
            throw new Exception("Failed to create sample images.");

        // Create a new Word document and insert the GIF image twice.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with GIF images:");
        builder.InsertImage(gifPath);
        builder.Writeln();
        builder.InsertImage(gifPath);
        string originalDocPath = Path.Combine(artifactsDir, "original.docx");
        doc.Save(originalDocPath);

        // Mapping from GIF to PNG (in this simple case it's a single entry).
        Dictionary<string, string> gifToPngMap = new Dictionary<string, string>
        {
            { gifPath, pngPath }
        };

        // Replace all GIF images in the document with their PNG equivalents.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                // Use the mapping to find the replacement PNG file.
                // Here we assume all GIFs use the same sample PNG.
                string replacementPng = pngPath;
                shape.ImageData.SetImage(replacementPng);
            }
        }

        // Save the modified document.
        string modifiedDocPath = Path.Combine(artifactsDir, "modified.docx");
        doc.Save(modifiedDocPath);

        // Validate that the output file exists.
        if (!File.Exists(modifiedDocPath))
            throw new Exception("Modified document was not saved.");

        // Optional validation: ensure no GIF images remain.
        NodeCollection finalShapes = doc.GetChildNodes(NodeType.Shape, true);
        int remainingGifCount = 0;
        foreach (Shape shape in finalShapes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
                remainingGifCount++;
        }
        if (remainingGifCount > 0)
            throw new Exception("Some GIF images were not replaced.");

        // Indicate successful completion (no interactive prompts).
        Console.WriteLine("GIF images successfully replaced with PNG equivalents.");
    }
}
