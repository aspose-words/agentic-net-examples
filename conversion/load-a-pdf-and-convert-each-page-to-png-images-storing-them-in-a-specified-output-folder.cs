using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Paths for the sample PDF and the output folder for PNG images.
        string pdfPath = "sample.pdf";
        string outputFolder = "OutputImages";

        // Ensure the output directory exists.
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Create a sample PDF document with three pages.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Sample PDF page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Sample PDF page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Sample PDF page 3.");

        // Save the document as PDF.
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // -----------------------------------------------------------------
        // Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // Convert each page of the PDF to a separate PNG image.
        // -----------------------------------------------------------------
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Configure image save options for PNG format.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page.
                PageSet = new PageSet(pageIndex),

                // Set the resolution (dpi). Higher values produce higher‑quality images.
                Resolution = 300
                // ImageSize is omitted because it requires System.Drawing.Size,
                // which is prohibited by the Aspose.Drawing rules.
            };

            // Build the output file name.
            string outputPath = Path.Combine(outputFolder, $"page_{pageIndex + 1}.png");

            // Save the current page as PNG.
            pdfDoc.Save(outputPath, options);

            // Verify that the PNG file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create PNG for page {pageIndex + 1}.");
        }

        Console.WriteLine("PDF conversion to PNG images completed.");
    }
}
