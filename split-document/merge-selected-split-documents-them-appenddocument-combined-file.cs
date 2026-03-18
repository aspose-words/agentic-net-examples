using System;
using Aspose.Words;

class Program
{
    static void Main(string[] args)
    {
        // Example usage: merge two documents.
        // Replace the file names with the actual paths of your split documents.
        string[] inputFiles = { "doc1.docx", "doc2.docx" };
        string outputFile = "merged.docx";

        DocumentMerger.MergeDocuments(inputFiles, outputFile);
        Console.WriteLine($"Merged {inputFiles.Length} documents into '{outputFile}'.");
    }
}

public class DocumentMerger
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
