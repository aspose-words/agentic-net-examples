using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the sample PDF and the output folder.
        string inputPdfPath = "sample.pdf";
        string outputFolder = "output_images";

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF document with two pages.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is the content of the first page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the content of the second page.");
        sourceDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(inputPdfPath))
            throw new InvalidOperationException($"The input PDF '{inputPdfPath}' was not created.");

        // -----------------------------------------------------------------
        // Step 2: Prepare the output directory.
        // -----------------------------------------------------------------
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Step 3: Load the PDF and convert each page to a PNG image.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(inputPdfPath);
        int pageCount = pdfDoc.PageCount;

        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
        {
            // Configure image save options for PNG format and single page.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Save only the current page.
                PageSet = new PageSet(pageIndex)
            };

            // Build the output file name.
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the specific page as PNG.
            pdfDoc.Save(outputPath, options);

            // Verify that the PNG file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1} at '{outputPath}'.");
        }

        // All pages have been converted successfully.
    }
}
