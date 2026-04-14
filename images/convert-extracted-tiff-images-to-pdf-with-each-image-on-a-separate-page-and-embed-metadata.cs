using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ConvertTiffToPdf
{
    public static void Main()
    {
        // Directory for generated files
        string outputDir = Directory.GetCurrentDirectory();

        // Create sample TIFF images (3 pages)
        int imageCount = 3;
        string[] tiffFiles = new string[imageCount];
        for (int i = 0; i < imageCount; i++)
        {
            string fileName = Path.Combine(outputDir, $"sample{i + 1}.tiff");
            CreateSampleTiff(fileName, i);
            if (!File.Exists(fileName))
                throw new Exception($"Failed to create TIFF file: {fileName}");
            tiffFiles[i] = fileName;
        }

        // Create a new Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each TIFF image on a separate page
        for (int i = 0; i < imageCount; i++)
        {
            // Insert image and obtain the shape
            Shape shape = builder.InsertImage(tiffFiles[i]);

            // Validate that the shape actually contains an image
            if (!shape.HasImage)
                throw new Exception($"Inserted shape does not contain an image for file {tiffFiles[i]}");

            // Add a page break after each image except the last one
            if (i < imageCount - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Embed custom metadata (custom document properties)
        doc.CustomDocumentProperties.Add("ImageCount", imageCount);
        for (int i = 0; i < imageCount; i++)
        {
            doc.CustomDocumentProperties.Add($"Image{i + 1}Path", tiffFiles[i]);
        }

        // Save the document as PDF with custom properties exported
        string pdfPath = Path.Combine(outputDir, "Result.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export custom properties as standard entries in the PDF Info dictionary
            CustomPropertiesExport = PdfCustomPropertiesExport.Standard
        };
        doc.Save(pdfPath, pdfOptions);

        // Validate PDF creation
        if (!File.Exists(pdfPath))
            throw new Exception("PDF file was not created.");

        // Optional: clean up generated TIFF files (comment out if you need to keep them)
        //foreach (var file in tiffFiles) File.Delete(file);
    }

    // Helper method to create a deterministic single‑page TIFF image
    private static void CreateSampleTiff(string filePath, int index)
    {
        int width = 200;
        int height = 200;

        // Create bitmap and graphics objects using Aspose.Drawing
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            // Fill background with white
            graphics.Clear(Aspose.Drawing.Color.White);

            // Draw a colored rectangle to differentiate images
            Aspose.Drawing.Color rectColor = Aspose.Drawing.Color.FromArgb(255, (index + 1) * 80, 0, 0);
            using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(rectColor))
            {
                graphics.FillRectangle(brush, 0, 0, width, height);
            }

            // Save as TIFF
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Tiff);
        }
    }
}
