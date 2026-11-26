using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with multiple pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 5)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        // Save the document as PDF (input for conversion).
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Determine even-numbered pages.
        List<int> evenPages = new List<int>();
        for (int page = 1; page <= pdfDoc.PageCount; page++)
        {
            if (page % 2 == 0)
                evenPages.Add(page);
        }

        // Export each even page as a separate PNG file.
        foreach (int pageNumber in evenPages)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(pageNumber)
            };

            string outputPath = $"page_{pageNumber}.png";
            pdfDoc.Save(outputPath, options);

            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"PNG file for page {pageNumber} was not created.");
        }

        // Optional: clean up generated files (comment out if you want to keep them).
        // File.Delete(pdfPath);
        // foreach (int pageNumber in evenPages)
        // {
        //     File.Delete($"page_{pageNumber}.png");
        // }
    }
}
