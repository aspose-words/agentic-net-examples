using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input documents and output PDFs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");
        string assetsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Assets");

        // Ensure folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);
        Directory.CreateDirectory(assetsFolder);

        // -----------------------------------------------------------------
        // Create a sample corporate logo image (PNG) without using System.Drawing.
        // The image is a 100x100 red square encoded as PNG bytes.
        // -----------------------------------------------------------------
        byte[] logoBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAIAAAD/gAIDAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJ" +
            "bWFnZVJlYWR5ccllPAAAAAlwSFlzAAALEwAACxMBAJqcGAAAABl0RVh0Q3JlYXRpb24gVGltZQAw" +
            "Mi8wOS8xOeZ0ZgAAABl0RVh0U291cmNlAEFzcG9zZSBXb3JkcyB2Mi4wLjB6cK6UAAAAB3RJTUUH" +
            "5QYVFAwzM7c2VQAAAB1pVFh0Q29tbWVudAAAAAAAV2l0aCBjb3Jwb3JhdGUgbG9nbyBpbWFnZSBm" +
            "b3IgdGVzdGluZy4AAAAASUVORK5CYII=");
        string logoPath = Path.Combine(assetsFolder, "logo.png");
        File.WriteAllBytes(logoPath, logoBytes);

        // -----------------------------------------------------------------
        // Create a few sample DOCX files to demonstrate batch processing.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"SampleDocument{i}.docx");
            // Create a blank document.
            Document doc = new Document();
            // Add some simple content.
            var builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document #{i}.");
            // Save the DOCX file.
            doc.Save(docPath, SaveFormat.Docx);
        }

        // -----------------------------------------------------------------
        // Batch convert each DOCX to PDF while applying the image watermark.
        // -----------------------------------------------------------------
        foreach (string docFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the source DOCX.
            Document document = new Document(docFile);

            // Configure watermark options (optional).
            ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
            {
                // Scale the image to 30% of its original size.
                Scale = 0.3,
                // Disable washout to keep original colors.
                IsWashout = false
            };

            // Apply the image watermark using the file path.
            document.Watermark.SetImage(logoPath, watermarkOptions);

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save as PDF.
            document.Save(pdfPath, SaveFormat.Pdf);
        }

        // -----------------------------------------------------------------
        // Simple verification: list generated PDF files.
        // -----------------------------------------------------------------
        Console.WriteLine("Batch conversion completed. Generated PDF files:");
        foreach (string pdf in Directory.GetFiles(outputFolder, "*.pdf"))
        {
            Console.WriteLine(pdf);
        }
    }
}
