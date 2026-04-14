using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define paths for the input folder and the output PDF.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.pdf");

        // Ensure the input folder exists.
        Directory.CreateDirectory(inputFolder);

        // Create sample DOCX files in the input folder.
        const int sampleCount = 3;
        for (int i = 1; i <= sampleCount; i++)
        {
            string filePath = Path.Combine(inputFolder, $"Sample{i}.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is the content of sample document {i}.");
            sampleDoc.Save(filePath, SaveFormat.Docx);
        }

        // Get all DOCX files from the folder.
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");

        // Load the first document as the master; then append the rest.
        Document? masterDoc = null;
        foreach (string file in docxFiles)
        {
            Document srcDoc = new Document(file);
            if (masterDoc == null)
            {
                // Use the first document as the base.
                masterDoc = srcDoc;
            }
            else
            {
                // Append subsequent documents, preserving their formatting.
                masterDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            }
        }

        // If no source files were found, create an empty document to avoid null reference.
        masterDoc ??= new Document();

        // Save the merged document as PDF.
        masterDoc.Save(outputPdf, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The merged PDF was not created.");

        // Validation: ensure the merged document contains content from all source files.
        // The number of sections should equal the number of source documents.
        if (masterDoc.Sections.Count != docxFiles.Length)
            throw new InvalidOperationException("The merged document does not contain all source sections.");

        // Confirmation output.
        Console.WriteLine($"Merged {docxFiles.Length} documents into '{outputPdf}'.");
    }
}
