using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchRtfToPdfA1a
{
    public static void Main()
    {
        // Define folders for input RTF files and output PDF/A‑1a files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputRtf");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample RTF documents.
        for (int i = 1; i <= 3; i++)
        {
            string rtfPath = Path.Combine(inputFolder, $"Sample{i}.rtf");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample document #{i} for legal archiving.");
            // Save as RTF.
            sampleDoc.Save(rtfPath, SaveFormat.Rtf);
        }

        // Get all RTF files in the input folder.
        string[] rtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
        if (rtfFiles.Length == 0)
            throw new InvalidOperationException("No RTF files were found for conversion.");

        // Convert each RTF file to PDF/A‑1a.
        foreach (string rtfFile in rtfFiles)
        {
            // Load the RTF document.
            Document doc = new Document(rtfFile);

            // Configure PDF save options for PDF/A‑1a compliance.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a
            };

            // Determine the output PDF file path.
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF/A‑1a.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF file was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF file: {pdfPath}");
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
