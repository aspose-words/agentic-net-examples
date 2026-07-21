using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files and output PDF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample DOCX files.
        for (int i = 1; i <= 3; i++)
        {
            // Create a blank document and add some sample text.
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample content for document {i}.");

            // Save the document as DOCX using the provided pattern.
            string docxPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            sampleDoc.Save(docxPath, SaveFormat.Docx);
        }

        // Process each DOCX file: add a company‑wide header and convert to PDF.
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docxFile in docxFiles)
        {
            // Load the existing DOCX file.
            Document doc = new Document(docxFile);

            // Insert a primary header with the company text.
            DocumentBuilder headerBuilder = new DocumentBuilder(doc);
            headerBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            headerBuilder.Writeln("Company Confidential");

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the provided pattern.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF conversion failed for '{docxFile}'.");
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch processing completed successfully.");
    }
}
