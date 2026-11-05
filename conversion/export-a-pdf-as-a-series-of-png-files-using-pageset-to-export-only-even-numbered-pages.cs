using System;
using System.IO;
using System.Linq;
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
                // Insert a page break to start a new page.
                builder.InsertBreak(BreakType.PageBreak);
            }
        }

        // Save the document as PDF (input for the conversion).
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The source PDF was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Determine even‑numbered pages (1‑based indexing).
        int[] evenPages = Enumerable.Range(1, pdfDoc.PageCount)
                                    .Where(p => p % 2 == 0)
                                    .ToArray();

        if (evenPages.Length == 0)
            throw new InvalidOperationException("The document does not contain any even‑numbered pages.");

        // Export each even page as a separate PNG file.
        const string pngPattern = "page_{0}.png";

        foreach (int pageNumber in evenPages)
        {
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Restrict saving to the current even page.
                PageSet = new PageSet(pageNumber)
            };

            string fileName = string.Format(pngPattern, pageNumber);
            pdfDoc.Save(fileName, saveOptions);

            if (!File.Exists(fileName))
                throw new InvalidOperationException($"Expected PNG file '{fileName}' was not created.");
        }

        // All done.
    }
}
