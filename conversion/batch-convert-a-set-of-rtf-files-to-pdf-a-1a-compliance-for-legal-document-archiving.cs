using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input RTF files and output PDF/A-1a files.
        string inputFolder = "InputRtf";
        string outputFolder = "OutputPdf";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample RTF files.
        for (int i = 1; i <= 3; i++)
        {
            // Create a new empty document.
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample RTF content for document {i}.");
            string inputPath = Path.Combine(inputFolder, $"sample{i}.rtf");
            // Save the document as RTF.
            sampleDoc.Save(inputPath, SaveFormat.Rtf);
        }

        // Get all RTF files from the input folder.
        string[] rtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
        if (rtfFiles.Length == 0)
        {
            throw new InvalidOperationException("No RTF files were found for conversion.");
        }

        // Prepare PDF/A-1a save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a
        };

        // Convert each RTF file to PDF/A-1a.
        foreach (string rtfFile in rtfFiles)
        {
            // Load the RTF document.
            Document doc = new Document(rtfFile);

            // Determine the output PDF file path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(rtfFile);
            string outputPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.pdf");

            // Save the document as PDF/A-1a.
            doc.Save(outputPath, pdfOptions);

            // Verify that the PDF file was created.
            if (!File.Exists(outputPath))
            {
                throw new InvalidOperationException($"Expected output PDF/A file was not created: {outputPath}");
            }
        }

        // Optional: indicate successful completion.
        Console.WriteLine("Batch conversion to PDF/A-1a completed successfully.");
    }
}
