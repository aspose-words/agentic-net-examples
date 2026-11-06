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
        // Prepare a folder for sample TIFF images.
        string inputFolder = "InputImages";
        Directory.CreateDirectory(inputFolder);

        // Create sample TIFF images.
        for (int i = 1; i <= 2; i++)
        {
            string imagePath = Path.Combine(inputFolder, $"image{i}.tif");
            using (Bitmap bitmap = new Bitmap(300, 300))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // Fill background with a deterministic color.
                    g.Clear(Color.FromArgb(255, 200 + i * 20, 200 - i * 20, 255));
                }
                // Save as TIFF.
                bitmap.Save(imagePath, ImageFormat.Tiff);
            }
        }

        // Verify that at least one TIFF image was created.
        string[] tiffFiles = Directory.GetFiles(inputFolder, "*.tif");
        if (tiffFiles.Length == 0)
            throw new Exception("No TIFF images were created.");

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each TIFF image on a separate page.
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            if (i > 0)
                builder.InsertBreak(BreakType.PageBreak);

            // Insert the image; let Aspose.Words handle scaling.
            builder.InsertImage(tiffFiles[i]);
        }

        // Embed metadata into the document.
        doc.BuiltInDocumentProperties.Title = "TIFF to PDF conversion";
        doc.BuiltInDocumentProperties.Author = "Aspose.Words Example";
        doc.BuiltInDocumentProperties.Subject = "Converted PDF";
        doc.BuiltInDocumentProperties.Keywords = "TIFF,PDF,Example";

        // Add a custom property indicating the source folder.
        doc.CustomDocumentProperties.Add("SourceFolder", inputFolder);

        // Save the document as PDF.
        string outputPdf = "output.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new Exception("PDF output was not created.");

        // Example completed successfully.
    }
}
