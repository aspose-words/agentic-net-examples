using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input RTF files and output PDF/A‑1a files.
        string inputFolder = "InputRtf";
        string outputFolder = "OutputPdf";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // Create sample RTF documents (the task requires local sample input).
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            // Create a blank document.
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Add deterministic content.
            builder.Writeln($"Sample RTF document #{i}");
            builder.Writeln("Legal text for archiving purposes.");
            builder.Writeln($"Generated on {DateTime.UtcNow:u}");

            // Save as RTF in the input folder.
            string rtfPath = Path.Combine(inputFolder, $"Sample{i}.rtf");
            sampleDoc.Save(rtfPath, SaveFormat.Rtf);
        }

        // ---------------------------------------------------------------
        // Batch convert each RTF file to PDF/A‑1a compliant PDF document.
        // ---------------------------------------------------------------
        string[] rtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
        int convertedCount = 0;

        foreach (string rtfFile in rtfFiles)
        {
            // Load the RTF document.
            Document doc = new Document(rtfFile);

            // Configure PDF save options for PDF/A‑1a compliance.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a
            };

            // Determine output PDF file path (same name, .pdf extension).
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF/A‑1a.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF file was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF file: {pdfPath}");

            convertedCount++;
        }

        // Simple confirmation (no interactive input required).
        Console.WriteLine($"Batch conversion completed. {convertedCount} file(s) converted to PDF/A‑1a.");
    }
}
