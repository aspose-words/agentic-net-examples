using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing types for colors

public class Program
{
    public static void Main()
    {
        // Prepare working folders
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample multi‑page Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Page {i}");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document as PDF (the source PDF)
        string pdfPath = Path.Combine(workDir, "Sample.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF back
        Document pdfDoc = new Document(pdfPath);

        // Convert each PDF page to a PNG image with a transparent background
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render with a transparent background using Aspose.Drawing.Color.
                // ImageSaveOptions.PaperColor expects System.Drawing.Color, so convert explicitly.
                PaperColor = System.Drawing.Color.FromArgb(Aspose.Drawing.Color.Transparent.ToArgb()),
                // Select the page to render
                PageSet = new PageSet(pageIndex)
            };

            string pngPath = Path.Combine(outputDir, $"Page_{pageIndex + 1}.png");
            pdfDoc.Save(pngPath, options);

            // Verify that the image file was created
            if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1}.");
        }

        // Final verification: at least one PNG should exist
        string[] pngFiles = Directory.GetFiles(outputDir, "*.png");
        if (pngFiles.Length == 0)
            throw new InvalidOperationException("No PNG images were generated.");

        // Optional cleanup (commented out)
        // Directory.Delete(workDir, true);
    }
}
