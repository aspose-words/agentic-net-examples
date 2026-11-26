using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for sample images.
        string imagesFolder = "SampleImages";
        Directory.CreateDirectory(imagesFolder);

        // Create two deterministic TIFF images.
        string[] tiffPaths = new string[2];
        for (int i = 0; i < tiffPaths.Length; i++)
        {
            string filePath = Path.Combine(imagesFolder, $"image{i + 1}.tif");
            CreateSampleTiff(filePath, $"Sample {i + 1}");
            tiffPaths[i] = filePath;
        }

        // Validate that TIFF files were created.
        foreach (string path in tiffPaths)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"Failed to create TIFF image: {path}");
        }

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each TIFF on a separate page.
        for (int i = 0; i < tiffPaths.Length; i++)
        {
            if (i > 0)
                builder.InsertBreak(BreakType.PageBreak);

            // Insert the image.
            builder.InsertImage(tiffPaths[i]);
        }

        // Embed metadata into the document.
        doc.BuiltInDocumentProperties.Title = "Converted TIFF Images to PDF";
        doc.BuiltInDocumentProperties.Author = "Aspose.Words Example";
        doc.BuiltInDocumentProperties.Subject = "TIFF to PDF conversion with metadata";
        doc.BuiltInDocumentProperties.Keywords = "Aspose.Words, TIFF, PDF, metadata";

        // Add a custom document property.
        doc.CustomDocumentProperties.Add("SourceFolder", imagesFolder);

        // Save the document as PDF.
        string outputPdf = "ConvertedImages.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException($"PDF file was not created: {outputPdf}");
    }

    // Helper method to create a deterministic TIFF image with simple text.
    private static void CreateSampleTiff(string filePath, string text)
    {
        int width = 400;
        int height = 300;

        // Use explicit Aspose.Drawing types to avoid ambiguity.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);

        // Clear background.
        graphics.Clear(Aspose.Drawing.Color.White);

        // Draw a rectangle border.
        graphics.DrawRectangle(Aspose.Drawing.Pens.Black, 0, 0, width - 1, height - 1);

        // Draw the supplied text.
        using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
        {
            graphics.DrawString(text, font, Aspose.Drawing.Brushes.Black, new Aspose.Drawing.PointF(10, 10));
        }

        // Save as TIFF.
        bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Tiff);

        // Clean up resources.
        graphics.Dispose();
        bitmap.Dispose();
    }
}
