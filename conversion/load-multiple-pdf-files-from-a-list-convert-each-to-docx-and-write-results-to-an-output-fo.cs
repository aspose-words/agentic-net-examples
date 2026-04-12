using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base working directory.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample PDF files.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample PDF document number {i}.");
            string pdfPath = Path.Combine(inputDir, $"Sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Get all PDF files from the input folder.
        string[] pdfFiles = Directory.GetFiles(inputDir, "*.pdf");

        foreach (string pdfFilePath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDocument = new Document(pdfFilePath);

            // Determine the output DOCX path.
            string outputFileName = Path.GetFileNameWithoutExtension(pdfFilePath) + ".docx";
            string outputPath = Path.Combine(outputDir, outputFileName);

            // Convert and save as DOCX.
            pdfDocument.Save(outputPath, SaveFormat.Docx);

            // Verify that the output file was created.
            if (!File.Exists(outputPath))
            {
                throw new InvalidOperationException($"Failed to create output file: {outputPath}");
            }
        }

        // Optional: indicate completion.
        Console.WriteLine("PDF to DOCX conversion completed successfully.");
    }
}
