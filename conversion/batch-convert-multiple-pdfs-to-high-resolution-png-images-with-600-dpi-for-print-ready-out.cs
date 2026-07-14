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

        // Ensure clean start.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files.
        const int samplePdfCount = 3;
        for (int i = 1; i <= samplePdfCount; i++)
        {
            // Create a simple Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample PDF #{i}");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Second page of PDF #{i}");
            // Save as PDF.
            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Batch convert each PDF to high‑resolution PNG images (600 DPI).
        foreach (string pdfFile in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Prepare image save options for PNG with 600 DPI.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = 600 // Set both horizontal and vertical resolution.
            };

            // Render each page separately.
            for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
            {
                // Specify which page to render.
                pngOptions.PageSet = new PageSet(pageIndex);

                // Build output file name: <pdfName>_page<1‑based>.png
                string pdfName = Path.GetFileNameWithoutExtension(pdfFile);
                string pngPath = Path.Combine(outputFolder, $"{pdfName}_page{pageIndex + 1}.png");

                // Save the rendered page as PNG.
                pdfDoc.Save(pngPath, pngOptions);

                // Validate that the PNG file was created.
                if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
                    throw new InvalidOperationException($"Failed to create PNG: {pngPath}");
            }
        }

        // All conversions completed successfully.
        Console.WriteLine("Batch conversion completed. PNG files are located in:");
        Console.WriteLine(outputFolder);
    }
}
