using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Folder that contains the DOCX files to be merged (created in the current directory)
        string sourceFolder = Path.Combine(Directory.GetCurrentDirectory(), "SourceDocs");
        Directory.CreateDirectory(sourceFolder);

        // Ensure there is at least one DOCX file to process; create sample files if none exist
        string[] existingDocs = Directory.GetFiles(sourceFolder, "*.docx");
        if (existingDocs.Length == 0)
        {
            for (int i = 1; i <= 2; i++)
            {
                Document sample = new Document();
                DocumentBuilder builder = new DocumentBuilder(sample);
                builder.Writeln($"This is sample document {i}.");
                string samplePath = Path.Combine(sourceFolder, $"Sample{i}.docx");
                sample.Save(samplePath);
            }
        }

        // Path for the final merged PDF document (saved in the current directory)
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.pdf");

        // Create an empty master document
        Document masterDoc = new Document();

        // Retrieve all DOCX files in the specified folder (non‑recursive)
        string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx");

        foreach (string filePath in docxFiles)
        {
            // Load each source document
            Document srcDoc = new Document(filePath);

            // Append the source document to the master while keeping its original formatting
            masterDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the combined document as PDF; format is inferred from the .pdf extension
        masterDoc.Save(outputPdfPath);

        Console.WriteLine($"Merged PDF created at: {outputPdfPath}");
    }
}
