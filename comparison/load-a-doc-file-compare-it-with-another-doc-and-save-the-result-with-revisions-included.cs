using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");
        builderOriginal.Writeln("This is the original document.");

        // Create the revised document with differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello world!"); // Modified line.
        builderRevised.Writeln("This is the revised document with an extra line.");

        // Ensure both documents have no revisions before comparison.
        if (original.HasRevisions || revised.HasRevisions)
        {
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");
        }

        // Compare the documents. The original document will receive revisions describing the changes.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the comparison result (original document now contains revisions).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparedResult.docx");
        original.Save(outputPath);

        // Optionally, write the revision count to the console for verification.
        Console.WriteLine($"Comparison completed. Revisions count: {original.Revisions.Count}");
        Console.WriteLine($"Result saved to: {outputPath}");
    }
}
