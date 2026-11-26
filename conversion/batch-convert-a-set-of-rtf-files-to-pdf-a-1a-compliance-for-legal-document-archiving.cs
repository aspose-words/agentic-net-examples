using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchRtfToPdfA
{
    public static void Main()
    {
        // Define folders for input RTF files and output PDF/A-1a files.
        string inputFolder = "InputRtf";
        string outputFolder = "OutputPdf";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample RTF documents.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample RTF content for file {i}.");

            string rtfPath = Path.Combine(inputFolder, $"sample{i}.rtf");
            sampleDoc.Save(rtfPath, SaveFormat.Rtf);
        }

        // Get all RTF files in the input folder.
        string[] rtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
        if (rtfFiles.Length == 0)
        {
            throw new InvalidOperationException("No RTF files were found for conversion.");
        }

        // Convert each RTF file to PDF/A-1a.
        foreach (string rtfFilePath in rtfFiles)
        {
            // Load the RTF document.
            Document rtfDoc = new Document(rtfFilePath);

            // Configure PDF/A-1a compliance options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a
            };

            // Determine the output PDF file path.
            string outputPdfPath = Path.Combine(
                outputFolder,
                Path.GetFileNameWithoutExtension(rtfFilePath) + ".pdf");

            // Save the document as PDF/A-1a.
            rtfDoc.Save(outputPdfPath, pdfOptions);

            // Verify that the PDF file was created.
            if (!File.Exists(outputPdfPath))
            {
                throw new InvalidOperationException($"Failed to create PDF/A-1a file: {outputPdfPath}");
            }
        }

        // Optional: indicate successful completion.
        Console.WriteLine("Batch conversion to PDF/A-1a completed successfully.");
    }
}
