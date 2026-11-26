using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output folders.
        string inputFolder = "input_pdfs";
        string outputFolder = "output_docx";

        // Ensure a clean environment.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files.
        List<string> pdfPaths = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            CreateSamplePdf(pdfPath, $"This is sample PDF number {i}.");
            pdfPaths.Add(pdfPath);
        }

        // Convert each PDF to DOCX.
        foreach (string pdfPath in pdfPaths)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfPath);

            // Determine the output DOCX file path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfPath);
            string docxPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}.docx");

            // Save the document as DOCX.
            pdfDoc.Save(docxPath, SaveFormat.Docx);

            // Verify that the DOCX file was created.
            if (!File.Exists(docxPath))
                throw new InvalidOperationException($"Conversion failed: {docxPath} was not created.");
        }
    }

    private static void CreateSamplePdf(string path, string text)
    {
        // Create a simple document with one paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);

        // Save the document as PDF.
        doc.Save(path, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample PDF at {path}.");
    }
}
