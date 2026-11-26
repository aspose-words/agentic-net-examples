using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample background image.
        const string backgroundImagePath = "bg.png";
        CreateSampleImage(backgroundImagePath, 600, 800);

        // Create a DOCX document and add the background image as a shape behind the text.
        const string docPath = "sample.docx";
        CreateDocumentWithBackground(docPath, backgroundImagePath);

        // Extract all images (including background) from the document and save as PNG files.
        ExtractImages(docPath);
    }

    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightBlue);
                // Draw a simple rectangle for visual distinction.
                using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5))
                {
                    graphics.DrawRectangle(pen, 50, 50, width - 100, height - 100);
                }
            }
            bitmap.Save(filePath);
        }

        if (!File.Exists(filePath))
            throw new Exception($"Failed to create sample image: {filePath}");
    }

    private static void CreateDocumentWithBackground(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a shape that covers the whole page and set it behind the text.
        Shape backgroundShape = new Shape(doc, ShapeType.Image);
        backgroundShape.ImageData.SetImage(imagePath);
        backgroundShape.Width = doc.FirstSection.PageSetup.PageWidth;
        backgroundShape.Height = doc.FirstSection.PageSetup.PageHeight;
        backgroundShape.WrapType = WrapType.None;
        backgroundShape.BehindText = true;

        // Insert the shape at the beginning of the document.
        builder.InsertNode(backgroundShape);

        // Add some sample text to demonstrate that the image is behind the text.
        builder.Writeln("This document contains a background image.");

        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new Exception($"Failed to create document: {docPath}");
    }

    private static void ExtractImages(string docPath)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outputPath = $"extracted-{extractedCount}.png";
                shape.ImageData.Save(outputPath);
                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new Exception("No images were extracted from the document.");

        // Validation: ensure all extracted files exist.
        for (int i = 0; i < extractedCount; i++)
        {
            string path = $"extracted-{i}.png";
            if (!File.Exists(path))
                throw new Exception($"Expected extracted image not found: {path}");
        }
    }
}
