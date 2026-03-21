using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Determine a root folder relative to the executable.
        string rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(rootFolder);

        // If there are no DOCX files, create a couple of sample documents.
        string[] existingDocs = Directory.GetFiles(rootFolder, "*.docx", SearchOption.TopDirectoryOnly);
        if (existingDocs.Length == 0)
        {
            CreateSampleDocument(Path.Combine(rootFolder, "Sample1.docx"), "First sample document.");
            CreateSampleDocument(Path.Combine(rootFolder, "Sample2.docx"), "Second sample document.");
        }

        // Create an empty master document.
        Document masterDoc = new Document();

        // Retrieve all DOCX files from the root folder and its subfolders.
        string[] docxFiles = Directory.GetFiles(rootFolder, "*.docx", SearchOption.AllDirectories);

        // Append each document to the master using destination styles.
        foreach (string filePath in docxFiles)
        {
            Document srcDoc = new Document(filePath);
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
        }

        // Ensure the output folder exists.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Export the combined document to PDF.
        string outputPath = Path.Combine(outputFolder, "MasterDocument.pdf");
        masterDoc.Save(outputPath);

        Console.WriteLine($"Combined document saved to: {outputPath}");
    }

    private static void CreateSampleDocument(string path, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(path);
    }
}
