using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main(string[] args)
    {
        // Create two temporary Word documents with sample content.
        string[] inputFiles = CreateSampleDocuments();

        // Define the output merged document path.
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "merged.docx");

        // Merge the documents.
        DocumentMerger.MergeDocuments(inputFiles, outputFile);
        Console.WriteLine($"Merged {inputFiles.Length} documents into '{outputFile}'.");

        // Clean up the temporary source files.
        foreach (var file in inputFiles)
        {
            try { File.Delete(file); } catch { /* ignore */ }
        }
    }

    private static string[] CreateSampleDocuments()
    {
        string[] paths = new string[2];
        for (int i = 0; i < 2; i++)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some sample text.
            builder.Writeln($"This is the content of document {i + 1}.");

            // Save to a temporary .docx file.
            string tempPath = Path.Combine(Path.GetTempPath(), $"sample_doc_{i + 1}_{Guid.NewGuid()}.docx");
            doc.Save(tempPath);
            paths[i] = tempPath;
        }
        return paths;
    }
}

public static class DocumentMerger
{
    /// <summary>
    /// Merges the specified Word documents into a single document.
    /// </summary>
    /// <param name="inputFiles">Array of file paths to the documents to be merged.</param>
    /// <param name="outputFile">File path for the combined document.</param>
    public static void MergeDocuments(string[] inputFiles, string outputFile)
    {
        // Create a blank destination document.
        Document dstDoc = new Document();

        // Load each source document and append it to the destination.
        foreach (string filePath in inputFiles)
        {
            Document srcDoc = new Document(filePath);
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the combined document to the specified output file.
        dstDoc.Save(outputFile);
    }
}
