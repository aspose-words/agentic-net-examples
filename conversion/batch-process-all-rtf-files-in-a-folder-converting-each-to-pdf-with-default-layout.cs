using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Base directory of the application.
        string baseDir = Directory.GetCurrentDirectory();

        // Input folder for RTF files and output folder for PDFs.
        string inputDir = Path.Combine(baseDir, "InputRtf");
        string outputDir = Path.Combine(baseDir, "OutputPdf");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample RTF documents.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample RTF content for file {i}.");

            string rtfPath = Path.Combine(inputDir, $"Sample{i}.rtf");
            // Save as RTF using the default layout.
            sampleDoc.Save(rtfPath, SaveFormat.Rtf);
        }

        // Batch convert each RTF file in the input folder to PDF.
        string[] rtfFiles = Directory.GetFiles(inputDir, "*.rtf");
        foreach (string rtfFile in rtfFiles)
        {
            // Load the RTF document.
            Document doc = new Document(rtfFile);

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfFile) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);

            // Save as PDF with default layout.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Expected PDF was not created: {pdfPath}");
        }

        // Optional: indicate successful completion.
        Console.WriteLine($"Converted {rtfFiles.Length} RTF file(s) to PDF in '{outputDir}'.");
    }
}
