using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create sample TIFF images.
        const int imageCount = 3;
        const int width = 600;
        const int height = 800;
        string[] tiffFiles = new string[imageCount];

        for (int i = 0; i < imageCount; i++)
        {
            string fileName = $"sample{i + 1}.tif";
            using (Bitmap bitmap = new Bitmap(width, height))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // Fill background with a distinct color.
                    Aspose.Drawing.Color bgColor = i % 2 == 0 ? Aspose.Drawing.Color.LightBlue : Aspose.Drawing.Color.LightGreen;
                    g.Clear(bgColor);
                }
                bitmap.Save(fileName);
            }
            tiffFiles[i] = Path.GetFullPath(fileName);
        }

        // Step 2: Create a new Word document that will be saved as PDF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure each image occupies a full page.
        double pageWidth = builder.PageSetup.PageWidth;
        double pageHeight = builder.PageSetup.PageHeight;

        for (int i = 0; i < tiffFiles.Length; i++)
        {
            // Insert the TIFF image scaled to page size.
            builder.InsertImage(tiffFiles[i], pageWidth, pageHeight);

            // Add a page break after each image except the last one.
            if (i < tiffFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Step 3: Save the document as PDF.
        string outputPdf = "output.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validation: ensure the PDF was created.
        if (!File.Exists(outputPdf) || new FileInfo(outputPdf).Length == 0)
            throw new Exception("Failed to create the PDF output.");

        // Cleanup: optional removal of temporary TIFF files.
        foreach (var file in tiffFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }
    }
}
