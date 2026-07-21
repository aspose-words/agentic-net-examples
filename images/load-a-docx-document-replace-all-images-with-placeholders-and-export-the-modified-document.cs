using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for sample images and documents
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        string placeholderImagePath = Path.Combine(artifactsDir, "placeholder.png");
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");

        // Create a sample image (used as original content)
        CreateSampleImage(sampleImagePath, 100, 100, Aspose.Drawing.Color.LightGray);

        // Create a placeholder image (will replace original images)
        CreateSampleImage(placeholderImagePath, 100, 100, Aspose.Drawing.Color.DarkGray);

        // Build a document that contains a few images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        builder.InsertBreak(BreakType.PageBreak);
        builder.InsertImage(sampleImagePath);
        doc.Save(inputDocPath);

        // Load the document we just created
        Document loadedDoc = new Document(inputDocPath);

        // Replace every image with the placeholder image
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Ensure the shape actually contains an image before replacement
                shape.ImageData.SetImage(placeholderImagePath);
            }
        }

        // Save the modified document
        loadedDoc.Save(outputDocPath);

        // Validate that the output file was created
        if (!File.Exists(outputDocPath))
            throw new Exception("The output document was not created.");

        // Optional: indicate success (no interactive prompts)
        Console.WriteLine("Document processing completed successfully.");
    }

    // Helper method to create a deterministic bitmap and save it to a file
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            // Save as PNG to ensure compatibility
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
