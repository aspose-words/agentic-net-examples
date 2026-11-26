using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Thumbnails");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files
        for (int i = 1; i <= 2; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample PDF content for file {i}.");
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create sample PDF: {pdfPath}");
        }

        // Process each PDF and generate a JPEG thumbnail of the first page
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            Document pdfDocument = new Document(pdfFile);

            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Save only the first page (zero‑based index)
                PageSet = new PageSet(0),
                // Low quality JPEG (0‑100)
                JpegQuality = 30
            };

            string thumbnailPath = Path.Combine(outputFolder,
                Path.GetFileNameWithoutExtension(pdfFile) + ".jpg");
            pdfDocument.Save(thumbnailPath, jpegOptions);

            if (!File.Exists(thumbnailPath))
                throw new InvalidOperationException($"Thumbnail was not created for {pdfFile}");
        }

        // Indicate success
        Console.WriteLine("PDF thumbnail generation completed successfully.");
    }
}
