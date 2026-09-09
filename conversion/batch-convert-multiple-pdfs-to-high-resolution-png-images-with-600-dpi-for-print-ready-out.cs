using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders for input PDFs and output PNGs.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputPdfs");
        string outputDir = Path.Combine(baseDir, "OutputPngs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample PDF files to act as the batch source.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            builder.Writeln($"Sample PDF {i} - Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample PDF {i} - Page 2.");

            string pdfPath = Path.Combine(inputDir, $"Sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create sample PDF: {pdfPath}");
        }

        // Batch convert each PDF to high‑resolution PNG images (600 DPI), one image per page.
        foreach (string pdfFile in Directory.GetFiles(inputDir, "*.pdf"))
        {
            Document pdfDoc = new Document(pdfFile);

            for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
            {
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
                {
                    // Render only the current page.
                    PageSet = new PageSet(pageIndex),
                    // Set the required resolution for print‑ready output.
                    Resolution = 600
                };

                string pngFileName = $"{Path.GetFileNameWithoutExtension(pdfFile)}_page{pageIndex + 1}.png";
                string pngPath = Path.Combine(outputDir, pngFileName);

                pdfDoc.Save(pngPath, options);

                if (!File.Exists(pngPath))
                    throw new InvalidOperationException($"Failed to create PNG image: {pngPath}");
            }
        }

        // Indicate successful completion (no interactive prompts).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
