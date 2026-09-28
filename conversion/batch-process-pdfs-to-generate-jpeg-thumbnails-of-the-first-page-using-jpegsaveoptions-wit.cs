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
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample PDF document {i}");
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Batch process PDFs to generate JPEG thumbnails of the first page
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            Document doc = new Document(pdfFile);

            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                PageSet = new PageSet(0), // first page (zero‑based)
                JpegQuality = 30          // low quality
            };

            string thumbnailPath = Path.Combine(outputFolder,
                Path.GetFileNameWithoutExtension(pdfFile) + ".jpg");
            doc.Save(thumbnailPath, jpegOptions);

            if (!File.Exists(thumbnailPath) || new FileInfo(thumbnailPath).Length == 0)
            {
                throw new InvalidOperationException($"Thumbnail was not created for '{pdfFile}'.");
            }
        }

        // Verification output (non‑interactive)
        Console.WriteLine($"Processed {pdfFiles.Length} PDF(s). Thumbnails saved to '{outputFolder}'.");
    }
}
