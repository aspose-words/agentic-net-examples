using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file with two pages.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is the content of page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the content of page 2.");
        string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The sample PDF file was not created.");

        // Load the PDF document.
        Document pdfDocument = new Document(pdfPath);

        // Export each page as a separate PNG image.
        for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(pageIndex)
            };

            string outputImagePath = $"page_{pageIndex + 1}.png";
            pdfDocument.Save(outputImagePath, options);

            if (!File.Exists(outputImagePath))
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1}.");
        }

        // Optional: Clean up the sample PDF if not needed.
        // File.Delete(pdfPath);
    }
}
