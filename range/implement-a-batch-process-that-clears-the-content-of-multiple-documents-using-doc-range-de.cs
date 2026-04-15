using System;
using System.IO;
using Aspose.Words;

public class BatchClearDocuments
{
    public static void Main()
    {
        // Directories for source and cleared documents.
        string baseDir = Directory.GetCurrentDirectory();
        string sourceDir = Path.Combine(baseDir, "SourceDocs");
        string outputDir = Path.Combine(baseDir, "ClearedDocs");

        // Ensure the directories exist.
        Directory.CreateDirectory(sourceDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample documents with some text.
        for (int i = 1; i <= 3; i++)
        {
            string srcPath = Path.Combine(sourceDir, $"Document{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample content for document {i}.");
            doc.Save(srcPath);
        }

        // Batch process: load each document, clear its content, and save the result.
        foreach (string srcPath in Directory.GetFiles(sourceDir, "*.docx"))
        {
            Document doc = new Document(srcPath);

            // Delete all characters in the document's range.
            doc.Range.Delete();

            // Save the cleared document to the output folder.
            string fileName = Path.GetFileName(srcPath);
            string destPath = Path.Combine(outputDir, $"Cleared_{fileName}");
            doc.Save(destPath);
        }
    }
}
