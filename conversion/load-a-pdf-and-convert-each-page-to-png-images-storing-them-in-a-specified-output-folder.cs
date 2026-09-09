using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class PdfToPngConverter
{
    public static void Main()
    {
        // Define paths for the sample PDF and the output folder.
        string inputPdfPath = "input.pdf";
        string outputFolder = "output_images";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Create a sample PDF document.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Add some content spanning multiple pages.
        builder.Writeln("This is page 1 of the sample PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2 of the sample PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3 of the sample PDF.");

        // Save the document as a PDF file.
        sampleDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(inputPdfPath))
            throw new InvalidOperationException("The input PDF file was not created.");

        // -----------------------------------------------------------------
        // Load the PDF and convert each page to a separate PNG image.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(inputPdfPath);

        for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
        {
            // Configure image save options for PNG format.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page.
                PageSet = new PageSet(pageIndex)
            };

            // Build the output file name for the current page.
            string outputFilePath = Path.Combine(outputFolder, $"page_{pageIndex + 1}.png");

            // Save the page as a PNG image.
            pdfDocument.Save(outputFilePath, pngOptions);

            // Validate that the image file was created.
            if (!File.Exists(outputFilePath))
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1}.");
        }

        // All pages have been converted successfully.
        Console.WriteLine("PDF conversion to PNG completed successfully.");
    }
}
