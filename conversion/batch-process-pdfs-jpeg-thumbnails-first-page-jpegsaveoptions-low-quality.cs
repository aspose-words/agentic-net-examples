using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfThumbnailBatch
{
    static void Main()
    {
        // Folder containing source PDF files (relative to the executable)
        string inputFolder = Path.Combine(AppContext.BaseDirectory, "InputPdfs");
        string outputFolder = Path.Combine(AppContext.BaseDirectory, "Thumbnails");

        // Ensure the directories exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Get PDF files; if none, inform the user and exit gracefully
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        if (pdfFiles.Length == 0)
        {
            Console.WriteLine($"No PDF files found in '{inputFolder}'. Place PDFs there and rerun.");
            return;
        }

        foreach (string pdfPath in pdfFiles)
        {
            Document doc = new Document(pdfPath);

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                PageSet = new PageSet(0),
                JpegQuality = 10
            };

            string thumbFileName = Path.GetFileNameWithoutExtension(pdfPath) + "_thumb.jpg";
            string thumbPath = Path.Combine(outputFolder, thumbFileName);

            doc.Save(thumbPath, options);
            Console.WriteLine($"Created thumbnail: {thumbPath}");
        }
    }
}
