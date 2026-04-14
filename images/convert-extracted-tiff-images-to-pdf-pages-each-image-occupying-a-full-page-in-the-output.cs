using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Layout;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directory for generated files
        string outputDir = Directory.GetCurrentDirectory();

        // Create sample TIFF images
        string[] tiffFiles = { Path.Combine(outputDir, "sample1.tiff"), Path.Combine(outputDir, "sample2.tiff") };
        CreateSampleTiff(tiffFiles[0], Aspose.Drawing.Color.LightBlue);
        CreateSampleTiff(tiffFiles[1], Aspose.Drawing.Color.LightGreen);

        // Create a new document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each TIFF image as a full‑page picture
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            // Insert the image; InsertImage returns the created Shape
            Shape imgShape = builder.InsertImage(tiffFiles[i]);

            // Ensure the shape actually contains an image
            if (!imgShape.HasImage)
                throw new InvalidOperationException($"Image not loaded from {tiffFiles[i]}.");

            // Resize the shape to fill the page
            imgShape.Width = doc.FirstSection.PageSetup.PageWidth;
            imgShape.Height = doc.FirstSection.PageSetup.PageHeight;

            // Position the image on the page
            imgShape.WrapType = WrapType.None;
            imgShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            imgShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Add a page break after each image except the last one
            if (i < tiffFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF
        string pdfPath = Path.Combine(outputDir, "ImagesToPdf.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the PDF file.");
    }

    // Creates a deterministic single‑page TIFF image with a solid background color
    private static void CreateSampleTiff(string filePath, Aspose.Drawing.Color background)
    {
        const int width = 600;
        const int height = 800;

        // Create bitmap and graphics objects
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);

        // Fill background
        graphics.Clear(background);

        // Save as TIFF
        bitmap.Save(filePath, ImageFormat.Tiff);

        // Clean up
        graphics.Dispose();
        bitmap.Dispose();
    }
}
