using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input folder and sample PDF files.
        string inputFolder = "InputPdfs";
        Directory.CreateDirectory(inputFolder);

        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample PDF document #{i}");
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create sample PDF: {pdfPath}");
        }

        // Prepare output folder for PNG images.
        string outputFolder = "OutputPngs";
        Directory.CreateDirectory(outputFolder);

        // Batch convert each PDF to high‑resolution PNG images (600 DPI), one image per page.
        foreach (string pdfFile in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            Document pdfDocument = new Document(pdfFile);
            int pageCount = pdfDocument.PageCount;

            for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                string outputFileName = $"{Path.GetFileNameWithoutExtension(pdfFile)}_page{pageIndex + 1}.png";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
                {
                    PageSet = new PageSet(pageIndex),
                    Resolution = 600
                };

                pdfDocument.Save(outputPath, options);

                if (!File.Exists(outputPath))
                    throw new InvalidOperationException($"Failed to create PNG image: {outputPath}");
            }
        }

        // Verify that at least one PNG file was created.
        string[] createdPngs = Directory.GetFiles(outputFolder, "*.png");
        if (createdPngs.Length == 0)
            throw new InvalidOperationException("No PNG images were generated.");

        // Example completed successfully.
    }
}
