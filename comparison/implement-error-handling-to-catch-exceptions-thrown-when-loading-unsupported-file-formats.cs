using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create first sample document.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("First version of the document.");

        // Create second sample document with a difference.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Second version with a change.");

        // Save the sample documents to the current directory.
        string currentDir = Directory.GetCurrentDirectory();
        string path1 = Path.Combine(currentDir, "doc1.docx");
        string path2 = Path.Combine(currentDir, "doc2.docx");
        doc1.Save(path1);
        doc2.Save(path2);

        // Create an unsupported file (plain text) to trigger an exception.
        string unsupportedPath = Path.Combine(currentDir, "unsupported.txt");
        File.WriteAllText(unsupportedPath, "Just some plain text.");

        // Attempt to load the unsupported file and handle the exception.
        try
        {
            Document unsupportedDoc = new Document(unsupportedPath);
            // The line above is expected to throw; if it doesn't, we simply ignore the document.
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Caught exception while loading unsupported file: {ex.Message}");
        }

        // Load the supported documents for comparison.
        Document baseDoc = new Document(path1);
        Document revisedDoc = new Document(path2);

        // Perform the comparison.
        baseDoc.Compare(revisedDoc, "Comparer", DateTime.Now);

        // Verify that revisions were generated.
        if (baseDoc.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the comparison result.
        string resultPath = Path.Combine(currentDir, "comparisonResult.docx");
        baseDoc.Save(resultPath);
        Console.WriteLine($"Comparison completed. Result saved to: {resultPath}");
    }
}
