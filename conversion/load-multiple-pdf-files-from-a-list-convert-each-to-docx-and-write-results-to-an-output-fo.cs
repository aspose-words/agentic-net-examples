using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output DOCX files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocx");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample PDF content {i}");
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Get all PDF files from the input folder.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");

        // Convert each PDF to DOCX and save to the output folder.
        foreach (string pdfFilePath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDocument = new Document(pdfFilePath);

            // Determine the output DOCX path.
            string docxFileName = Path.GetFileNameWithoutExtension(pdfFilePath) + ".docx";
            string docxPath = Path.Combine(outputFolder, docxFileName);

            // Save as DOCX.
            pdfDocument.Save(docxPath, SaveFormat.Docx);

            // Verify that the DOCX file was created.
            if (!File.Exists(docxPath))
                throw new InvalidOperationException($"Expected output DOCX was not created: {docxPath}");
        }

        // Optional: indicate completion (no interactive prompts).
        Console.WriteLine("PDF to DOCX conversion completed successfully.");
    }
}
