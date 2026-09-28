using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample background image.
        const string backgroundImagePath = "bg.png";
        CreateSampleImage(backgroundImagePath, 800, 600);

        // Create a DOCX document and add the image as a background-like shape.
        const string docPath = "sample.docx";
        CreateDocumentWithBackgroundImage(docPath, backgroundImagePath);

        // Extract images from the document's background settings (shape images) and save as PNG files.
        ExtractBackgroundImages(docPath);
    }

    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
                // Draw a simple rectangle for visual distinction.
                graphics.DrawRectangle(new Pen(Color.DarkBlue, 5), 50, 50, width - 100, height - 100);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }

        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image at '{filePath}'.");
    }

    private static void CreateDocumentWithBackgroundImage(string docPath, string imagePath)
    {
        // Initialize a new empty document.
        Document doc = new Document();

        // Ensure the image file exists.
        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        // Create a shape that covers the whole page and place it behind the text.
        Shape backgroundShape = new Shape(doc, ShapeType.Image);
        backgroundShape.ImageData.SetImage(imagePath);
        backgroundShape.Width = doc.FirstSection.PageSetup.PageWidth;
        backgroundShape.Height = doc.FirstSection.PageSetup.PageHeight;
        backgroundShape.WrapType = WrapType.None;
        backgroundShape.BehindText = true;

        // Append the shape to the first paragraph of the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(backgroundShape);

        // Save the document.
        doc.Save(docPath);

        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to save document at '{docPath}'.");
    }

    private static void ExtractBackgroundImages(string docPath)
    {
        // Load the document.
        Document doc = new Document(docPath);

        // Collect all shape nodes.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        for (int i = 0; i < shapeNodes.Count; i++)
        {
            Shape shape = (Shape)shapeNodes[i];
            if (shape.HasImage)
            {
                string outputFileName = $"background-{extractedCount + 1}.png";
                shape.ImageData.Save(outputFileName);
                extractedCount++;

                if (!File.Exists(outputFileName))
                    throw new InvalidOperationException($"Failed to save extracted image '{outputFileName}'.");
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No background images were extracted from the document.");
    }
}
