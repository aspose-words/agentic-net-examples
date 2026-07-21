using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output PNGs.
        string inputFolder = "InputPdfs";
        string outputFolder = "OutputPngs";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample PDF files to demonstrate batch conversion.
        for (int i = 1; i <= 3; i++)
        {
            // Create a new blank document.
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Add some content and a page break for multiple pages.
            builder.Writeln($"Sample PDF document #{i}");
            builder.Writeln("This is a line of text on the first page.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is a line of text on the second page.");

            // Save the document as PDF in the input folder.
            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create sample PDF: {pdfPath}");
        }

        // Process each PDF file in the input folder.
        foreach (string pdfFile in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Convert each page of the PDF to a high‑resolution PNG image.
            for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
            {
                // Prepare the output PNG file name.
                string pngFileName = $"{Path.GetFileNameWithoutExtension(pdfFile)}_Page{pageIndex + 1}.png";
                string pngPath = Path.Combine(outputFolder, pngFileName);

                // Configure image save options: PNG format, 600 DPI, render only the current page.
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
                {
                    Resolution = 600,
                    PageSet = new PageSet(pageIndex)
                };

                // Save the page as a PNG image.
                pdfDoc.Save(pngPath, options);

                // Verify that the PNG image was created.
                if (!File.Exists(pngPath))
                    throw new InvalidOperationException($"Failed to create PNG image: {pngPath}");
            }
        }

        // All conversions completed successfully.
        Console.WriteLine("Batch conversion of PDFs to high‑resolution PNGs completed.");
    }
}
