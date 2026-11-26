using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output PNGs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPngs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample PDF content for file {i}.");
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create sample PDF: {pdfPath}");
        }

        // Batch convert each PDF to high‑resolution PNG images (600 DPI).
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDocument = new Document(pdfFile);

            // Configure image save options for PNG with 600 DPI.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = 600
            };

            // Base name for output files.
            string baseFileName = Path.GetFileNameWithoutExtension(pdfFile);

            // Save each page as a separate PNG file.
            for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
            {
                // Restrict saving to the current page.
                saveOptions.PageSet = new PageSet(pageIndex);

                string pngPath = Path.Combine(outputFolder, $"{baseFileName}_Page_{pageIndex}.png");
                pdfDocument.Save(pngPath, saveOptions);

                // Validate that the PNG file was created and is not empty.
                if (!File.Exists(pngPath))
                    throw new InvalidOperationException($"Expected PNG file was not created: {pngPath}");

                FileInfo info = new FileInfo(pngPath);
                if (info.Length == 0)
                    throw new InvalidOperationException($"Created PNG file is empty: {pngPath}");
            }
        }
    }
}
