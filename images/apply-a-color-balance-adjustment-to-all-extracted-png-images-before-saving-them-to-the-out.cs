using System;
using System.IO;
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
        string outputDir = Path.Combine(artifactsDir, "Output");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample PNG image using Aspose.Drawing
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        const int imgWidth = 200;
        const int imgHeight = 200;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            // Fill background with white
            g.Clear(Color.White);

            // Draw a simple red rectangle
            using (SolidBrush brush = new SolidBrush(Color.Red))
            {
                g.FillRectangle(brush, 20, 20, imgWidth - 40, imgHeight - 40);
            }

            // Save the bitmap as PNG
            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the PNG image several times
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image three times
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(sampleImagePath);
            builder.Writeln(); // add a line break between images
        }

        // Save the document
        string docPath = Path.Combine(artifactsDir, "SampleDocument.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document, adjust color balance of each PNG image,
        //    and save the adjusted images to the output folder
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Apply a simple color‑balance‑like adjustment.
            // Here we modify brightness and contrast as a proxy for color balance.
            shape.ImageData.Brightness = 0.8f; // brighter
            shape.ImageData.Contrast = 0.6f;   // slightly higher contrast

            // Save the adjusted image
            string outFile = Path.Combine(outputDir, $"Extracted_{extractedCount}.png");
            shape.ImageData.Save(outFile);
            extractedCount++;
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one image was saved
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and saved.");

        // Optional: verify that files exist (throws if any missing)
        for (int i = 0; i < extractedCount; i++)
        {
            string filePath = Path.Combine(outputDir, $"Extracted_{i}.png");
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Expected output file not found: {filePath}");
        }

        // Program completes without requiring user interaction.
    }
}
