using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input documents and output PDF.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputPdfPath = Path.Combine(baseDir, "MergedOutput.pdf");

        // Ensure the input directory exists.
        Directory.CreateDirectory(inputDir);

        // Create sample DOCX files inside the input directory.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputDir, $"Doc{i}.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample document {i} content.");
            sampleDoc.Save(docPath, SaveFormat.Docx);
        }

        // Create an empty master document that will receive the appended content.
        Document masterDoc = new Document();

        // Append each DOCX file in the input directory to the master document.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document srcDoc = new Document(filePath);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged document as a PDF file.
        masterDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The merged PDF file was not created.");

        // Load the PDF back and verify that it contains text from all source documents.
        Document pdfDoc = new Document(outputPdfPath);
        string pdfText = pdfDoc.GetText();

        for (int i = 1; i <= 3; i++)
        {
            if (!pdfText.Contains($"Sample document {i} content."))
                throw new InvalidOperationException($"Merged PDF is missing content from Doc{i}.docx.");
        }

        // Optional: clean up sample files (comment out if you want to inspect them).
        //Directory.Delete(inputDir, true);
    }
}
