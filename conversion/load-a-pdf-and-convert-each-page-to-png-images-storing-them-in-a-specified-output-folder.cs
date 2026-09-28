using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Define paths.
        string inputPdfPath = "sample.pdf";
        string outputFolder = "output_images";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Create a sample PDF if it does not already exist.
        if (!File.Exists(inputPdfPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is the content of the first page.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is the content of the second page.");
            sampleDoc.Save(inputPdfPath, SaveFormat.Pdf);
        }

        // Load the PDF document.
        Document pdfDocument = new Document(inputPdfPath);

        // Get the number of pages in the PDF.
        int pageCount = pdfDocument.PageCount;

        // Convert each page to a separate PNG image.
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
        {
            string outputImagePath = Path.Combine(outputFolder, $"page_{pageIndex + 1}.png");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);
            // Save only the current page.
            saveOptions.PageSet = new PageSet(pageIndex);

            pdfDocument.Save(outputImagePath, saveOptions);

            // Validate that the image was created.
            if (!File.Exists(outputImagePath))
            {
                throw new InvalidOperationException($"Expected image file was not created: {outputImagePath}");
            }
        }
    }
}
