using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;

public class PdfBatchToDocxConverter
{
    public static void Main()
    {
        // Define folders for input PDFs and output DOCX files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocx");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files to demonstrate the conversion.
        CreateSamplePdfFiles(inputFolder, count: 3);

        // Get all PDF files from the input folder.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");

        // Convert each PDF to DOCX and save to the output folder.
        foreach (string pdfPath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDocument = new Document(pdfPath);

            // Determine the output DOCX file path.
            string docxFileName = Path.GetFileNameWithoutExtension(pdfPath) + ".docx";
            string docxPath = Path.Combine(outputFolder, docxFileName);

            // Save the document as DOCX.
            pdfDocument.Save(docxPath, SaveFormat.Docx);

            // Validate that the DOCX file was created.
            if (!File.Exists(docxPath))
                throw new InvalidOperationException($"Conversion failed: '{docxPath}' was not created.");
        }

        // Optional: indicate completion.
        Console.WriteLine("All PDF files have been converted to DOCX.");
    }

    // Helper method to create a specified number of sample PDF files.
    private static void CreateSamplePdfFiles(string folderPath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            // Create a new blank document.
            Document doc = new Document();

            // Add sample text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample PDF document number {i}.");

            // Define the PDF file path.
            string pdfPath = Path.Combine(folderPath, $"sample{i}.pdf");

            // Save the document as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF file was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create sample PDF: '{pdfPath}'.");
        }
    }
}
