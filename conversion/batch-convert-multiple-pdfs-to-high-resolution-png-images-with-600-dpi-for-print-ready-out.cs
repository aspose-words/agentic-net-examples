using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchPdfToPngConverter
{
    public static void Main()
    {
        // Define folders for input PDFs and output PNGs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPngs");

        // Ensure clean environment.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files.
        CreateSamplePdf(Path.Combine(inputFolder, "Sample1.pdf"), "First sample PDF.", 2);
        CreateSamplePdf(Path.Combine(inputFolder, "Sample2.pdf"), "Second sample PDF.", 3);

        // Process each PDF in the input folder.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfPath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDocument = new Document(pdfPath);

            // Convert each page to a separate high‑resolution PNG.
            for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
            {
                ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
                pngOptions.Resolution = 600; // 600 DPI for print‑ready quality.
                pngOptions.PageSet = new PageSet(pageIndex); // Render only the current page.

                string pngFileName = $"{Path.GetFileNameWithoutExtension(pdfPath)}_Page{pageIndex + 1}.png";
                string pngPath = Path.Combine(outputFolder, pngFileName);

                pdfDocument.Save(pngPath, pngOptions);

                // Validate that the PNG was created.
                if (!File.Exists(pngPath))
                    throw new InvalidOperationException($"Failed to create PNG file: {pngPath}");
            }
        }

        // Optional: write a simple completion message.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to create a sample PDF with the specified text and page count.
    private static void CreateSamplePdf(string filePath, string title, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < pageCount; i++)
        {
            builder.Writeln($"{title} - Page {i + 1}");
            if (i < pageCount - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(filePath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample PDF: {filePath}");
    }
}
