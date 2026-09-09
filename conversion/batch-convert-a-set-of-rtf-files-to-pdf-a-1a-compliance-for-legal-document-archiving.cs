using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Directories for input RTF files and output PDF/A‑1a files.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputRtf");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample RTF documents.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample RTF content #{i}");
            string rtfPath = Path.Combine(inputDir, $"sample{i}.rtf");
            sampleDoc.Save(rtfPath, SaveFormat.Rtf);
        }

        // Batch convert each RTF file to PDF/A‑1a.
        string[] rtfFiles = Directory.GetFiles(inputDir, "*.rtf");
        foreach (string rtfFile in rtfFiles)
        {
            // Load the RTF document.
            Document doc = new Document(rtfFile);

            // Configure PDF/A‑1a compliance.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a
            };

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfFile) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);

            // Save as PDF/A‑1a.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF file: {pdfPath}");
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
