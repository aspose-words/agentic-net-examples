using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths.
        string samplePdfPath = "sample.pdf";
        string outputFolder = "OutputImages";

        // Ensure the output folder exists.
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Create a sample PDF document (required because we cannot assume an
        // external file exists). The document will have two pages.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Sample PDF page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Sample PDF page 2.");
        sampleDoc.Save(samplePdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(samplePdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // -----------------------------------------------------------------
        // Load the PDF and convert each page to a separate PNG image.
        // -----------------------------------------------------------------
        Document pdfDocument = new Document(samplePdfPath);

        for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
        {
            // Configure image save options for PNG and select the current page.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(pageIndex),
                // Optional: increase resolution for higher quality output.
                Resolution = 300
            };

            string outputFile = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");
            pdfDocument.Save(outputFile, options);
        }

        // -----------------------------------------------------------------
        // Validate that each PNG file was created.
        // -----------------------------------------------------------------
        for (int i = 0; i < pdfDocument.PageCount; i++)
        {
            string expectedPath = Path.Combine(outputFolder, $"Page_{i + 1}.png");
            if (!File.Exists(expectedPath))
                throw new InvalidOperationException($"Expected image file was not created: {expectedPath}");
        }

        // All pages have been successfully converted.
        Console.WriteLine("PDF conversion to PNG completed successfully.");
    }
}
