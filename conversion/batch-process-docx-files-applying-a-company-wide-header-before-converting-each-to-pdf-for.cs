using System;
using System.IO;
using Aspose.Words;

public class BatchDocxToPdfWithHeader
{
    public static void Main()
    {
        // Define folders for input DOCX files and output PDF files.
        string inputFolder = "InputDocs";
        string outputFolder = "OutputPdfs";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files if the input folder is empty.
        CreateSampleDocuments(inputFolder);

        // Process each DOCX file in the input folder.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Add a company‑wide header.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Company Confidential");

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Convert and save the document as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF file was not created: {pdfPath}");
        }
    }

    // Creates a few sample DOCX files with simple content.
    private static void CreateSampleDocuments(string folder)
    {
        // If there are already DOCX files, skip creation.
        if (Directory.GetFiles(folder, "*.docx").Length > 0)
            return;

        for (int i = 1; i <= 3; i++)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {i}");
            builder.Writeln("This is a placeholder paragraph for testing batch conversion.");

            // Save the document as DOCX.
            string docxPath = Path.Combine(folder, $"Sample{i}.docx");
            doc.Save(docxPath, SaveFormat.Docx);
        }
    }
}
