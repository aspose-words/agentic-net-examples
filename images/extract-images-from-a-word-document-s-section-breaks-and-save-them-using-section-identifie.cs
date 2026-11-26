using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractImagesBySection
{
    public static void Main()
    {
        // Prepare deterministic file names.
        const string docPath = "sample.docx";
        const string img1Path = "sample1.png";
        const string img2Path = "sample2.png";

        // -------------------------------------------------
        // Step 1: Create sample images using Aspose.Drawing.
        // -------------------------------------------------
        CreateSampleImage(img1Path, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage(img2Path, Aspose.Drawing.Color.LightCoral);

        // -------------------------------------------------
        // Step 2: Build a Word document with two sections,
        // each containing one image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section – insert first image.
        builder.InsertImage(img1Path);
        // Insert a section break (new page) to start the second section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section – insert second image.
        builder.InsertImage(img2Path);

        // Save the document.
        doc.Save(docPath);

        // -------------------------------------------------
        // Step 3: Load the document and extract images per section.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        int totalExtracted = 0;

        // Iterate through each section.
        for (int secIndex = 0; secIndex < loadedDoc.Sections.Count; secIndex++)
        {
            Section section = loadedDoc.Sections[secIndex];
            // Find all Shape nodes that contain images within this section.
            NodeCollection shapes = section.Body.GetChildNodes(NodeType.Shape, true);
            int imgIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string outFile = $"section-{secIndex + 1}-image-{imgIndex + 1}.png";
                    shape.ImageData.Save(outFile);
                    imgIndex++;
                    totalExtracted++;
                }
            }
        }

        // Validation: ensure at least one image was extracted.
        if (totalExtracted == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Cleanup: optional – delete temporary files if desired.
        // File.Delete(img1Path);
        // File.Delete(img2Path);
    }

    // Helper method to create a deterministic PNG image.
    private static void CreateSampleImage(string filePath, Aspose.Drawing.Color fillColor)
    {
        const int width = 200;
        const int height = 100;
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
