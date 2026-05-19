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

        // Create a few sample RTF documents.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample RTF document #{i}");
            builder.Writeln("This is a legal document intended for archiving.");
            string rtfPath = Path.Combine(inputFolder, $"Sample{i}.rtf");
            sampleDoc.Save(rtfPath, SaveFormat.Rtf);
        }

        // Batch convert each RTF file to PDF/A‑1a.
        string[] rtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
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

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF for '{rtfFile}'.");
        }

        // All conversions completed successfully.
    }
}
