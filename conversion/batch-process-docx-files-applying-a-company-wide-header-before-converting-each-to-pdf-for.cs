using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Define folders for input DOCX files and output PDF files.
        string inputFolder = "InputDocs";
        string outputFolder = "OutputPdfs";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX documents to simulate existing files.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample content for document {i}.");
            string inputPath = Path.Combine(inputFolder, $"Sample{i}.docx");
            sampleDoc.Save(inputPath, SaveFormat.Docx);
        }

        // Process each DOCX file: add a header and convert to PDF.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Add a company‑wide header.
            DocumentBuilder headerBuilder = new DocumentBuilder(doc);
            headerBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            headerBuilder.Writeln("Company Header");

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Expected PDF was not created: {pdfPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch processing completed successfully.");
    }
}
