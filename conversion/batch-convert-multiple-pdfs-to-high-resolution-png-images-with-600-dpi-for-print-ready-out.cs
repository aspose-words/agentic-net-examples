using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchPdfToPngConverter
{
    public static void Main()
    {
        // Prepare folders.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPngs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files.
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample PDF content for document {i}.");
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create sample PDF: {pdfPath}");
        }

        // Batch convert each PDF to high‑resolution PNG (600 DPI).
        foreach (string pdfFile in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            Document pdfDoc = new Document(pdfFile);

            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = 600
            };

            string outputBaseName = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(pdfFile));
            string pngPath = outputBaseName + ".png";

            pdfDoc.Save(pngPath, pngOptions);

            if (!File.Exists(pngPath))
                throw new InvalidOperationException($"PNG conversion failed for: {pdfFile}");
        }
    }
}
