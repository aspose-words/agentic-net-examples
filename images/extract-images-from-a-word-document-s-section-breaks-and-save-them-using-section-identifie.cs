using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common provides Bitmap, Graphics, Color

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic sample image (100x100, solid blue)
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 100, 100, Aspose.Drawing.Color.Blue);

        // 2. Build a document with several sections, each containing the sample image
        string docPath = Path.Combine(artifactsDir, "DocumentWithSections.docx");
        BuildDocumentWithSections(docPath, sampleImagePath);

        // 3. Load the document and extract images per section
        ExtractImagesBySection(docPath, artifactsDir);
    }

    // Creates a PNG image using Aspose.Drawing and saves it to the specified path.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        var bitmap = new Bitmap(width, height);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(fillColor);
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Builds a document that contains three sections, each with the same image.
    private static void BuildDocumentWithSections(string docPath, string imagePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Section {i}");
            // Insert the sample image inline.
            builder.InsertImage(imagePath);
            // Add a section break after each section except the last.
            if (i < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        doc.Save(docPath);
    }

    // Extracts images from each section and saves them using a section‑based identifier.
    private static void ExtractImagesBySection(string docPath, string outputDir)
    {
        var doc = new Document(docPath);
        int totalSaved = 0;

        for (int secIndex = 0; secIndex < doc.Sections.Count; secIndex++)
        {
            Section section = doc.Sections[secIndex];
            // Collect all Shape nodes that contain images within this section.
            var shapes = section.GetChildNodes(NodeType.Shape, true)
                               .OfType<Shape>()
                               .Where(s => s.HasImage)
                               .ToList();

            if (!shapes.Any())
                throw new InvalidOperationException($"No images found in section {secIndex + 1}.");

            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = $"Section{secIndex + 1}_Image{imageIndex}{extension}";
                string fullPath = Path.Combine(outputDir, fileName);
                shape.ImageData.Save(fullPath);
                imageIndex++;
                totalSaved++;
            }
        }

        if (totalSaved == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        Console.WriteLine($"Extraction complete. {totalSaved} image(s) saved to \"{outputDir}\".");
    }
}
