using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing; // Required for the PaperColor property

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document.");
        builder.Writeln("Each page will be exported as a PNG image with a transparent background.");

        // Save the document as PDF.
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Export each page of the PDF to a separate PNG file with a transparent background.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render with a transparent background.
                PaperColor = Color.Transparent,
                // Select the current page.
                PageSet = new PageSet(pageIndex)
            };

            string pngPath = $"page_{pageIndex + 1}.png";
            pdfDoc.Save(pngPath, pngOptions);

            // Verify that the PNG file was created.
            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"Failed to create PNG file: {pngPath}");
        }

        // Optionally clean up the temporary PDF file.
        // File.Delete(pdfPath);
    }
}
