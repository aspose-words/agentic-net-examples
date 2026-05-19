using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // For Bitmap, Graphics, Color

public class ExtractImagesBySection
{
    public static void Main()
    {
        // Prepare a deterministic output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample PNG image that will be inserted into the document.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Build a sample document containing two sections, each with the same image.
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        BuildDocumentWithSections(docPath, sampleImagePath);

        // Load the document and extract images per section.
        Document doc = new Document(docPath);

        int totalExtracted = 0;
        for (int secIndex = 0; secIndex < doc.Sections.Count; secIndex++)
        {
            Section section = doc.Sections[secIndex];

            // Retrieve all Shape nodes inside this section (including nested ones).
            NodeCollection shapeNodes = section.GetChildNodes(NodeType.Shape, true);
            var imageShapes = shapeNodes.OfType<Shape>().Where(s => s.HasImage).ToList();

            if (imageShapes.Count == 0)
                continue; // No images in this section.

            for (int imgIndex = 0; imgIndex < imageShapes.Count; imgIndex++)
            {
                Shape shape = imageShapes[imgIndex];
                string ext = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outFile = Path.Combine(
                    artifactsDir,
                    $"Section_{secIndex + 1}_Image_{imgIndex + 1}{ext}");

                // Save the image to the file system.
                shape.ImageData.Save(outFile);
                totalExtracted++;
                Console.WriteLine($"Saved image: {outFile}");
            }
        }

        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        Console.WriteLine($"Extraction complete. Total images saved: {totalExtracted}");
    }

    // Creates a simple solid‑color PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (var bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (var graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.CornflowerBlue);
            }
            bitmap.Save(filePath);
        }
    }

    // Builds a document with two sections, each containing the same sample image.
    private static void BuildDocumentWithSections(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section with an image.
        builder.Writeln("Section 1");
        builder.InsertImage(imagePath);
        // Insert a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section with an image.
        builder.Writeln("Section 2");
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);
    }
}
