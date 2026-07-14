using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputPdfs");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files
        CreateSampleDocx(Path.Combine(inputFolder, "Doc1.docx"), "First sample document content.");
        CreateSampleDocx(Path.Combine(inputFolder, "Doc2.docx"), "Second sample document content.");

        // Process each DOCX: add header and convert to PDF
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docxPath in docxFiles)
        {
            // Load the DOCX
            Document doc = new Document(docxPath);

            // Ensure a primary header exists and add company-wide header text
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Company Confidential – Header");

            // Determine output PDF path
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docxPath);
            string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

            // Save as PDF
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate output
            if (!File.Exists(pdfPath))
            {
                throw new InvalidOperationException($"PDF was not created for '{docxPath}'.");
            }

            FileInfo pdfInfo = new FileInfo(pdfPath);
            if (pdfInfo.Length == 0)
            {
                throw new InvalidOperationException($"PDF file '{pdfPath}' is empty.");
            }
        }

        // Optional: indicate completion (no interactive wait)
        Console.WriteLine("Batch processing completed successfully.");
    }

    private static void CreateSampleDocx(string path, string bodyText)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(bodyText);
        doc.Save(path, SaveFormat.Docx);

        if (!File.Exists(path))
        {
            throw new InvalidOperationException($"Failed to create sample DOCX at '{path}'.");
        }
    }
}
