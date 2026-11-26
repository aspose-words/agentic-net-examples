using System;
using System.IO;
using Aspose.Words;
using SkiaSharp;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOCX files
        for (int i = 1; i <= 2; i++)
        {
            string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            doc.Save(docPath);
        }

        // Create a sample logo image (1x1 red PNG)
        string logoPath = Path.Combine(baseDir, "logo.png");
        byte[] logoBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6ZcAAAAASUVORK5CYII=");
        File.WriteAllBytes(logoPath, logoBytes);

        // Load the image into an SKBitmap (required by Aspose.Words watermark API)
        using var bitmap = SKBitmap.Decode(logoPath);
        if (bitmap == null)
        {
            Console.WriteLine("Failed to load logo image.");
            return;
        }

        // Process each DOCX: add image watermark and convert to PDF
        foreach (string docFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            var document = new Document(docFile);
            document.Watermark.SetImage(bitmap);
            string pdfFileName = Path.GetFileNameWithoutExtension(docFile) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);
            document.Save(pdfPath, SaveFormat.Pdf);
        }

        // List generated PDFs
        foreach (string pdf in Directory.GetFiles(outputDir, "*.pdf"))
        {
            Console.WriteLine($"Generated PDF: {pdf}");
        }
    }
}
