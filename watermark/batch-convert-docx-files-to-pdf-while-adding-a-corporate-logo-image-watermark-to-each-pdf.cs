using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic logo image (PNG) from a Base64 string.
        string logoPath = Path.Combine(baseDir, "logo.png");
        if (!File.Exists(logoPath))
        {
            // A 1x1 pixel transparent PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAn0B9pVYVZcAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(logoPath, pngBytes);
        }

        // Create a few sample DOCX files if none exist.
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample document #{i}.");
                string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
                sampleDoc.Save(docPath);
            }
        }

        // Process each DOCX: add image watermark and convert to PDF.
        foreach (string docxFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the source document.
            Document doc = new Document(docxFile);

            // Configure image watermark options.
            ImageWatermarkOptions wmOptions = new ImageWatermarkOptions
            {
                IsWashout = false, // Keep original colors.
                Scale = 0 // Auto‑scale to fit page margins.
            };

            // Apply the image watermark using the logo file.
            doc.Watermark.SetImage(logoPath, wmOptions);

            // Determine output PDF path.
            string pdfFile = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(docxFile) + ".pdf");

            // Save as PDF.
            doc.Save(pdfFile, SaveFormat.Pdf);
        }

        Console.WriteLine("Batch conversion completed. PDFs are located in:");
        Console.WriteLine(outputDir);
    }
}
