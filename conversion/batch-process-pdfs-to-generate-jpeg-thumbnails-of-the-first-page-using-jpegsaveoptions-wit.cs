using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders for input PDFs and output thumbnails.
        string inputDir = "InputPdfs";
        string outputDir = "Thumbnails";
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample PDF files.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample PDF document {i}");
            string pdfPath = Path.Combine(inputDir, $"sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Process each PDF: render the first page to a low‑quality JPEG thumbnail.
        foreach (string pdfPath in Directory.GetFiles(inputDir, "*.pdf"))
        {
            Document pdfDoc = new Document(pdfPath);

            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render only the first page (zero‑based index).
                PageSet = new PageSet(0),
                // Low quality to increase compression.
                JpegQuality = 10
            };

            string thumbnailPath = Path.Combine(
                outputDir,
                Path.GetFileNameWithoutExtension(pdfPath) + ".jpg");

            pdfDoc.Save(thumbnailPath, jpegOptions);

            // Validate that the thumbnail was created.
            if (!File.Exists(thumbnailPath) || new FileInfo(thumbnailPath).Length == 0)
                throw new InvalidOperationException($"Failed to create thumbnail for '{pdfPath}'.");
        }

        // Indicate successful completion.
        Console.WriteLine("Thumbnails generated successfully.");
    }
}
