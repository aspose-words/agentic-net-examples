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
        string[] tiffFiles = { "image1.tif", "image2.tif" };
        int width = 500;
        int height = 700;
        Color[] colors = { Color.LightBlue, Color.LightCoral };

        for (int i = 0; i < tiffFiles.Length; i++)
        {
            using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
            {
                using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
                {
                    g.Clear(colors[i % colors.Length]);
                }
                bitmap.Save(tiffFiles[i], ImageFormat.Tiff);
            }

            // Ensure the file was created.
            if (!File.Exists(tiffFiles[i]))
                throw new Exception($"Failed to create TIFF file: {tiffFiles[i]}");
        }

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each TIFF image on a separate full page.
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            if (i > 0)
                builder.InsertBreak(BreakType.PageBreak); // Start a new page for subsequent images.

            // Get page dimensions.
            double pageWidth = builder.PageSetup.PageWidth;
            double pageHeight = builder.PageSetup.PageHeight;

            // Insert the image to fill the entire page.
            builder.InsertImage(
                tiffFiles[i],
                RelativeHorizontalPosition.Page,
                0,
                RelativeVerticalPosition.Page,
                0,
                pageWidth,
                pageHeight,
                WrapType.None);
        }

        // Save the document as PDF.
        string outputPath = "output.pdf";
        doc.Save(outputPath);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
            throw new Exception("PDF output was not created.");

        // Cleanup sample images (optional).
        foreach (string file in tiffFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }
    }
}
