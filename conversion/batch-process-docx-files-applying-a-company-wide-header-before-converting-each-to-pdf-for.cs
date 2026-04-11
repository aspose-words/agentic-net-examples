using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files and output PDF files.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputPdfs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample DOCX documents.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"SampleDocument{i}.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is the body of sample document {i}.");
            sampleDoc.Save(docPath, SaveFormat.Docx);
        }

        // Process each DOCX file: add a header and convert to PDF.
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docxFile in docxFiles)
        {
            // Load the DOCX document.
            Document doc = new Document(docxFile);

            // Add a company‑wide header.
            DocumentBuilder headerBuilder = new DocumentBuilder(doc);
            headerBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            headerBuilder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            headerBuilder.Write("Company Confidential – Header");

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
            {
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
            }
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("Batch processing completed successfully.");
    }
}
