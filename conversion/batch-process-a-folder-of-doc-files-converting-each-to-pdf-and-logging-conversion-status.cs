using System;
using System.IO;
using Aspose.Words;

public class BatchDocToPdfConverter
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = "InputDocs";
        string outputDir = "OutputPdfs";

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample DOC files in the input directory.
        CreateSampleDoc(Path.Combine(inputDir, "Sample1.doc"), "This is the first sample document.");
        CreateSampleDoc(Path.Combine(inputDir, "Sample2.doc"), "This is the second sample document.");
        CreateSampleDoc(Path.Combine(inputDir, "Sample3.doc"), "This is the third sample document.");

        // Process each DOC file in the input folder.
        string[] docFiles = Directory.GetFiles(inputDir, "*.doc");
        foreach (string docFilePath in docFiles)
        {
            // Load the DOC file.
            Document doc = new Document(docFilePath);

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docFilePath) + ".pdf";
            string pdfFilePath = Path.Combine(outputDir, pdfFileName);

            // Convert and save as PDF.
            doc.Save(pdfFilePath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfFilePath))
                throw new InvalidOperationException($"Conversion failed: PDF not created for '{docFilePath}'.");

            // Log conversion status.
            Console.WriteLine($"Converted '{docFilePath}' to '{pdfFilePath}'.");
        }

        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Helper method that follows the documented doc‑to‑pdf creation pattern.
    private static void CreateSampleDoc(string filePath, string content)
    {
        // Create a new blank document.
        Document source = new Document();

        // Add content using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln(content);

        // Save the document as a DOC file.
        source.Save(filePath, SaveFormat.Doc);

        // Verify that the DOC file was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample DOC file at '{filePath}'.");
    }
}
