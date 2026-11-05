using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create sample TIFF images.
        CreateSampleTiff("sample1.tif", 600, 800, Aspose.Drawing.Color.LightBlue);
        CreateSampleTiff("sample2.tif", 800, 600, Aspose.Drawing.Color.LightGreen);

        // Gather all TIFF files in the current directory.
        string[] tiffFiles = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.tif");
        if (tiffFiles.Length == 0)
            throw new Exception("No TIFF files found to convert.");

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each TIFF image on a separate full page.
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            if (i > 0)
                builder.InsertBreak(BreakType.PageBreak);

            // Calculate image size to fill the page (respecting margins).
            double pageWidth = builder.PageSetup.PageWidth - builder.PageSetup.LeftMargin - builder.PageSetup.RightMargin;
            double pageHeight = builder.PageSetup.PageHeight - builder.PageSetup.TopMargin - builder.PageSetup.BottomMargin;

            // Insert the image with the calculated dimensions.
            builder.InsertImage(tiffFiles[i], pageWidth, pageHeight);
        }

        // Save the document as PDF.
        string outputPath = "output.pdf";
        doc.Save(outputPath);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the PDF output file.");

        // Clean up sample images (optional).
        foreach (string file in tiffFiles)
            File.Delete(file);
    }

    private static void CreateSampleTiff(string fileName, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        // Create a bitmap and fill it with a solid color.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
            }

            // Save the bitmap as a TIFF image.
            bitmap.Save(fileName, ImageFormat.Tiff);
        }
    }
}
