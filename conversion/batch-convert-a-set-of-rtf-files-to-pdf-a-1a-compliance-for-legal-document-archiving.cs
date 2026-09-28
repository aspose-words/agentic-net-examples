using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchRtfToPdfA
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = "InputRtf";
        string outputDir = "OutputPdf";

        // Ensure directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample RTF files.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample RTF content for document {i}.");
            string rtfPath = Path.Combine(inputDir, $"Sample{i}.rtf");
            sampleDoc.Save(rtfPath, SaveFormat.Rtf);
        }

        // Batch convert each RTF to PDF/A‑1a.
        string[] rtfFiles = Directory.GetFiles(inputDir, "*.rtf");
        foreach (string rtfFile in rtfFiles)
        {
            Document doc = new Document(rtfFile);

            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a
            };

            string outputFileName = Path.GetFileNameWithoutExtension(rtfFile) + ".pdf";
            string pdfPath = Path.Combine(outputDir, outputFileName);
            doc.Save(pdfPath, pdfOptions);
        }

        // Validate that all PDFs were created.
        foreach (string rtfFile in rtfFiles)
        {
            string expectedPdf = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(rtfFile) + ".pdf");
            if (!File.Exists(expectedPdf))
                throw new InvalidOperationException($"Expected PDF/A file was not created: {expectedPdf}");
        }

        // Optionally, indicate success (no interactive output required).
        Console.WriteLine("Batch conversion to PDF/A‑1a completed successfully.");
    }
}
