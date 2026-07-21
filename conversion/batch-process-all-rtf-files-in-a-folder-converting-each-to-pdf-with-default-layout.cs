using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input RTF files and output PDFs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputRtf");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Seed the input folder with a few sample RTF documents.
        for (int i = 1; i <= 3; i++)
        {
            // Create a new blank document.
            Document source = new Document();
            DocumentBuilder builder = new DocumentBuilder(source);
            builder.Writeln($"Sample content for file {i}.");

            // Save the document as RTF using the default layout.
            string rtfPath = Path.Combine(inputFolder, $"Sample{i}.rtf");
            source.Save(rtfPath, SaveFormat.Rtf);
        }

        // Process each RTF file in the input folder.
        string[] rtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
        foreach (string rtfFile in rtfFiles)
        {
            // Load the RTF document.
            Document doc = new Document(rtfFile);

            // Determine the corresponding PDF output path.
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Convert and save the document as PDF with default layout.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Optional: indicate completion (no interactive input).
        Console.WriteLine($"Converted {rtfFiles.Length} RTF file(s) to PDF in '{outputFolder}'.");
    }
}
