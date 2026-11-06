using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a folder to hold sample source DOCX files.
        string sourceFolder = Path.Combine(Directory.GetCurrentDirectory(), "SourceDocs");
        Directory.CreateDirectory(sourceFolder);

        // Generate a few sample DOCX documents.
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(sourceFolder, $"Sample{i}.docx");
            CreateSampleDocx(filePath, $"This is the content of sample document {i}.");
        }

        // Get all DOCX files in the folder, ordered to ensure deterministic processing.
        string[] docPaths = Directory.GetFiles(sourceFolder, "*.docx")
                                     .OrderBy(p => p)
                                     .ToArray();

        if (docPaths.Length == 0)
            throw new InvalidOperationException("No source documents were found to merge.");

        // Load the first document as the master document.
        Document masterDoc = new Document(docPaths[0]);

        // Append the remaining documents.
        for (int i = 1; i < docPaths.Length; i++)
        {
            Document srcDoc = new Document(docPaths[i]);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged result as PDF.
        string outputPdf = Path.Combine(Directory.GetCurrentDirectory(), "MasterDocument.pdf");
        masterDoc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException($"Failed to create output PDF at '{outputPdf}'.");

        // Validate that the merged document contains the expected number of sections.
        int expectedSections = docPaths.Length; // One section per source document.
        if (masterDoc.Sections.Count != expectedSections)
            throw new InvalidOperationException(
                $"Merged document has {masterDoc.Sections.Count} sections; expected {expectedSections}.");

        // Optional: indicate success.
        Console.WriteLine($"Successfully merged {expectedSections} documents into PDF at '{outputPdf}'.");
    }

    // Helper method to create a simple DOCX file with given text.
    private static void CreateSampleDocx(string path, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(path, SaveFormat.Docx);
    }
}
