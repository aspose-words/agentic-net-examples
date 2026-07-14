using System;
using System.IO;
using Aspose.Words;

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
        CreateSampleRtf(Path.Combine(inputFolder, "Sample1.rtf"), "First sample RTF content.");
        CreateSampleRtf(Path.Combine(inputFolder, "Sample2.rtf"), "Second sample RTF content.");
        CreateSampleRtf(Path.Combine(inputFolder, "Sample3.rtf"), "Third sample RTF content.");

        // Process each RTF file in the input folder.
        string[] rtfFiles = Directory.GetFiles(inputFolder, "*.rtf");
        foreach (string rtfPath in rtfFiles)
        {
            // Load the RTF document.
            Document doc = new Document(rtfPath);

            // Determine the corresponding PDF file name.
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the default layout.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF conversion failed for '{rtfPath}'. Expected file '{pdfPath}' was not created.");
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch conversion of RTF files to PDF completed successfully.");
    }

    // Helper method to create a simple RTF document with specified text.
    private static void CreateSampleRtf(string filePath, string content)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the provided content.
        builder.Writeln(content);

        // Save the document as RTF.
        doc.Save(filePath, SaveFormat.Rtf);

        // Verify that the RTF file was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample RTF file at '{filePath}'.");
    }
}
