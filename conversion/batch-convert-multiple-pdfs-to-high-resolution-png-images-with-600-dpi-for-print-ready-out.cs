using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory for the demo.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "BatchConversionDemo");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure clean folders.
        if (Directory.Exists(baseDir))
            Directory.Delete(baseDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample PDF files.
        const int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Sample PDF document {i} - Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample PDF document {i} - Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample PDF document {i} - Page 3");

            string pdfPath = Path.Combine(inputDir, $"Sample{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Batch convert each PDF to high‑resolution PNG images (600 DPI).
        string[] pdfFiles = Directory.GetFiles(inputDir, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Convert each page.
            for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
            {
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
                {
                    Resolution = 600f,                     // 600 DPI for print‑ready output.
                    PageSet = new PageSet(pageIndex)       // Render only the current page.
                };

                string outputFileName = $"{Path.GetFileNameWithoutExtension(pdfFile)}_Page{pageIndex + 1}.png";
                string outputPath = Path.Combine(outputDir, outputFileName);

                pdfDoc.Save(outputPath, options);

                // Validate that the image was created.
                if (!File.Exists(outputPath))
                    throw new InvalidOperationException($"Failed to create image: {outputPath}");
            }
        }

        // Optional: indicate completion.
        Console.WriteLine($"Conversion completed. Images are located in: {outputDir}");
    }
}
