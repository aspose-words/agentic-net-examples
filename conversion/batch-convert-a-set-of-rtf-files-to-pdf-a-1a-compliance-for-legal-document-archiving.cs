using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchRtfToPdfA1a
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Directory.GetCurrentDirectory();

        // Input folder for RTF files.
        string inputFolder = Path.Combine(baseDir, "InputRtf");
        // Output folder for PDF/A‑1a files.
        string outputFolder = Path.Combine(baseDir, "OutputPdf");

        // Ensure folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample RTF documents.
        CreateSampleRtf(Path.Combine(inputFolder, "Document1.rtf"), "Legal Contract – Clause 1");
        CreateSampleRtf(Path.Combine(inputFolder, "Document2.rtf"), "Legal Contract – Clause 2");
        CreateSampleRtf(Path.Combine(inputFolder, "Document3.rtf"), "Legal Contract – Clause 3");

        // Prepare PDF save options for PDF/A‑1a compliance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a
        };

        // Process each RTF file in the input folder.
        string[] rtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
        foreach (string rtfPath in rtfFiles)
        {
            // Load the RTF document.
            Document doc = new Document(rtfPath);

            // Determine the output PDF file name.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(rtfPath);
            string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

            // Save as PDF/A‑1a.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
            {
                throw new InvalidOperationException($"Failed to create PDF file: {pdfPath}");
            }
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method to create a simple RTF document with given content.
    private static void CreateSampleRtf(string filePath, string content)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(content);
        // Save as RTF.
        doc.Save(filePath, SaveFormat.Rtf);
        // Verify creation.
        if (!File.Exists(filePath))
        {
            throw new InvalidOperationException($"Failed to create RTF file: {filePath}");
        }
    }
}
