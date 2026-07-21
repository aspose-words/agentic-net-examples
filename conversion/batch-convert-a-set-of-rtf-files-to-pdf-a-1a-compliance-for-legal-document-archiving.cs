using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Directories for input RTF files and output PDF/A‑1a files.
        string inputDir = "InputRtf";
        string outputDir = "OutputPdf";

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample RTF documents.
        for (int i = 1; i <= 3; i++)
        {
            string rtfPath = Path.Combine(inputDir, $"Sample{i}.rtf");

            // Create a blank document and add sample text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i} for legal archiving.");

            // Save the document as RTF.
            doc.Save(rtfPath, SaveFormat.Rtf);
        }

        // Get all RTF files from the input directory.
        string[] rtfFiles = Directory.GetFiles(inputDir, "*.rtf");

        foreach (string rtfFile in rtfFiles)
        {
            // Load the RTF document.
            Document rtfDoc = new Document(rtfFile);

            // Configure PDF save options for PDF/A‑1a compliance.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a
            };

            // Determine the output PDF file path.
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfFile) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);

            // Save the document as PDF/A‑1a.
            rtfDoc.Save(pdfPath, pdfOptions);

            // Verify that the PDF file was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF file: {pdfPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
