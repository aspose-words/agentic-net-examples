using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input folder and sample PDF files
        string inputFolder = "InputPdfs";
        Directory.CreateDirectory(inputFolder);

        // Sample PDF 1
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Sample PDF content - Document 1.");
        string pdfPath1 = Path.Combine(inputFolder, "sample1.pdf");
        doc1.Save(pdfPath1, SaveFormat.Pdf);

        // Sample PDF 2
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Sample PDF content - Document 2.");
        string pdfPath2 = Path.Combine(inputFolder, "sample2.pdf");
        doc2.Save(pdfPath2, SaveFormat.Pdf);

        // Prepare output folder for thumbnails
        string outputFolder = "Thumbnails";
        Directory.CreateDirectory(outputFolder);

        // Batch process each PDF to generate a low‑quality JPEG thumbnail of the first page
        foreach (string pdfFile in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            Document pdfDocument = new Document(pdfFile);

            string outputFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".jpg";
            string jpegPath = Path.Combine(outputFolder, outputFileName);

            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                PageSet = new PageSet(0), // First page only
                JpegQuality = 30          // Low quality
            };

            pdfDocument.Save(jpegPath, jpegOptions);

            if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
            {
                throw new InvalidOperationException($"Failed to create thumbnail for '{pdfFile}'.");
            }
        }

        Console.WriteLine("Thumbnail generation completed successfully.");
    }
}
