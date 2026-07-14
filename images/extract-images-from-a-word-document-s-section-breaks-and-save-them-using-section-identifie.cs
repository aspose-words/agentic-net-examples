using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;

public class ExtractImagesBySection
{
    public static void Main()
    {
        // Prepare output directories
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string artifactsDir = Path.Combine(outputDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a deterministic sample image (sample.png)
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // Build a document with two sections, each containing an image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section image
        builder.InsertImage(sampleImagePath);

        // Insert a section break (new page) to start a new section
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section image
        builder.InsertImage(sampleImagePath);

        // Save the document
        string docPath = Path.Combine(artifactsDir, "DocumentWithSections.docx");
        doc.Save(docPath);

        // Load the document (optional, we can reuse the same instance)
        Document loadedDoc = new Document(docPath);

        // Extract images per section
        int sectionCount = loadedDoc.Sections.Count;
        for (int secIdx = 0; secIdx < sectionCount; secIdx++)
        {
            Section section = loadedDoc.Sections[secIdx];
            NodeCollection shapes = section.GetChildNodes(NodeType.Shape, true);
            int imageIdx = 0;

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string imageFileName = $"Section_{secIdx + 1}_Image_{imageIdx + 1}{extension}";
                    string imageFullPath = Path.Combine(artifactsDir, imageFileName);
                    shape.ImageData.Save(imageFullPath);
                    imageIdx++;
                }
            }

            // Validation: ensure at least one image was extracted for the section
            if (imageIdx == 0)
                throw new InvalidOperationException($"No images were found in section {secIdx + 1}.");
        }

        // Optional: indicate completion
        Console.WriteLine("Image extraction completed. Files are located in: " + artifactsDir);
    }

    // Helper method to create a simple white PNG image using Aspose.Drawing
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
