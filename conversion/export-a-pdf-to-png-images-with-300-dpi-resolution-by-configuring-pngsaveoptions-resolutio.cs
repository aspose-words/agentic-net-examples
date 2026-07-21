using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document and save it as PDF (input PDF).
        string pdfPath = "sample.pdf";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the source PDF.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Export each page of the PDF to a separate PNG image with 300 DPI resolution.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Configure image save options for PNG format.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page.
                PageSet = new PageSet(pageIndex),
                // Set both horizontal and vertical resolution to 300 DPI.
                Resolution = 300
            };

            string pngPath = $"page_{pageIndex + 1}.png";
            pdfDoc.Save(pngPath, options);

            // Verify the PNG image was created.
            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1}.");
        }

        // Clean up: optional deletion of the temporary PDF.
        // File.Delete(pdfPath);
    }
}
