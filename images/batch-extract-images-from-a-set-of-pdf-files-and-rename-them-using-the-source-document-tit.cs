using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class BatchPdfImageExtractor
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchImages");
        string inputFolder = Path.Combine(baseDir, "InputPdfs");
        string outputFolder = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample images that will be inserted into the PDFs.
        string sampleImage1 = Path.Combine(baseDir, "sample1.png");
        string sampleImage2 = Path.Combine(baseDir, "sample2.png");
        CreateSampleImage(sampleImage1, 200, 100, Aspose.Drawing.Color.LightBlue, "Img1");
        CreateSampleImage(sampleImage2, 150, 150, Aspose.Drawing.Color.LightGreen, "Img2");

        // Create sample PDF files with images and a title property.
        for (int i = 1; i <= 2; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Insert first image.
            builder.InsertImage(sampleImage1);
            builder.Writeln(); // Add a line break.
            // Insert second image.
            builder.InsertImage(sampleImage2);
            // Set the document title.
            doc.BuiltInDocumentProperties.Title = $"SampleDoc{i}";
            string pdfPath = Path.Combine(inputFolder, $"SampleDoc{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Process each PDF: extract images and rename them using the document title.
        foreach (string pdfPath in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            Document pdfDoc = new Document(pdfPath);
            string title = pdfDoc.BuiltInDocumentProperties.Title;
            if (string.IsNullOrEmpty(title))
                title = Path.GetFileNameWithoutExtension(pdfPath);

            NodeCollection shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 1;
            foreach (Shape shape in shapeNodes)
            {
                if (shape.HasImage)
                {
                    string outImagePath = Path.Combine(outputFolder, $"{title}_Image{imageIndex}.png");
                    shape.ImageData.Save(outImagePath);
                    imageIndex++;
                }
            }

            if (imageIndex == 1) // No images were saved.
                throw new InvalidOperationException($"No images were extracted from '{pdfPath}'.");
        }

        // Validation: ensure at least one image file exists in the output folder.
        if (Directory.GetFiles(outputFolder, "*.png").Length == 0)
            throw new InvalidOperationException("Image extraction completed but no output images were found.");

        // Example completed.
    }

    // Helper method to create a deterministic sample PNG image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backColor);
                // Simple drawing: fill a rectangle.
                using (var brush = new SolidBrush(Aspose.Drawing.Color.DarkBlue))
                {
                    graphics.FillRectangle(brush, 10, 10, width - 20, height - 20);
                }
                // Optionally, draw text if needed (omitted to avoid font ambiguity).
            }
            bitmap.Save(filePath);
        }
    }
}
