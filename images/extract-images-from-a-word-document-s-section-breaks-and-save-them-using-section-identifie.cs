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
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic sample image (sample.png).
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple rectangle to make the image recognizable.
                graphics.DrawRectangle(new Pen(Color.Black, 3), 20, 20, 160, 160);
            }
            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Build a sample document with several sections, each containing the image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        const int sectionCount = 3;
        for (int i = 1; i <= sectionCount; i++)
        {
            builder.Writeln($"This is content of section {i}.");
            // Insert the same sample image into each section.
            builder.InsertImage(sampleImagePath);

            // Add a section break after each section except the last one.
            if (i < sectionCount)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the document.
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract images from each section and save them using section identifiers.
        // -----------------------------------------------------------------
        int totalExtracted = 0;

        for (int secIndex = 0; secIndex < doc.Sections.Count; secIndex++)
        {
            Section section = doc.Sections[secIndex];
            // Get all Shape nodes inside the current section (including nested nodes).
            NodeCollection shapeNodes = section.GetChildNodes(NodeType.Shape, true);

            int imageInSection = 0;
            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // Determine the proper file extension based on the image type.
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string outFileName = $"Section_{secIndex + 1}_Image_{imageInSection}{extension}";
                    string outPath = Path.Combine(artifactsDir, outFileName);

                    // Save the image to the file system.
                    shape.ImageData.Save(outPath);
                    imageInSection++;
                    totalExtracted++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: output a simple summary.
        Console.WriteLine($"Extraction complete. Total images saved: {totalExtracted}");
        Console.WriteLine($"Files are located in: {artifactsDir}");
    }
}
