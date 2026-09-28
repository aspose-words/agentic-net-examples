using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with multiple pages.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Page 1 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content.");

        // Save the sample document as PDF.
        const string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The sample PDF was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Export each page as a separate PNG image.
        int pageCount = pdfDoc.PageCount;
        for (int i = 0; i < pageCount; i++)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(i) // Zero‑based page index.
            };

            string pngPath = $"page_{i + 1}.png";
            pdfDoc.Save(pngPath, options);

            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Expected PNG page {i + 1} was not created.");
        }

        // Optional: clean up created files (comment out if you want to inspect the outputs).
        // File.Delete(pdfPath);
        // for (int i = 1; i <= pageCount; i++) File.Delete($"page_{i}.png");
    }
}
