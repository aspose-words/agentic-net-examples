using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a sample PDF file.
        string inputPdfPath = "input.pdf";
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is the content of page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the content of page 2.");
        sampleDoc.Save(inputPdfPath, SaveFormat.Pdf);

        if (!File.Exists(inputPdfPath))
            throw new InvalidOperationException("The sample PDF was not created.");

        // Load the PDF document.
        Document pdfDocument = new Document(inputPdfPath);

        // Create output folder for PNG images.
        string outputFolder = "output_images";
        Directory.CreateDirectory(outputFolder);

        // Convert each page to a separate PNG image.
        int pageCount = pdfDocument.PageCount;
        for (int i = 0; i < pageCount; i++)
        {
            string outputPath = Path.Combine(outputFolder, $"page_{i + 1}.png");
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(i) // Zero‑based page index.
            };
            pdfDocument.Save(outputPath, options);

            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create image for page {i + 1}.");
        }
    }
}
