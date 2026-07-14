using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for the generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonDemo");
        Directory.CreateDirectory(outputDir);

        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world!");
        builderOriginal.Writeln("This is the original document.");

        // Create the revised document with differences.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello world!");
        builderRevised.Writeln("This is the revised document with changes.");

        // Perform the comparison. The original document will receive revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Output the number of revisions to the console (non‑interactive).
        Console.WriteLine($"Revisions detected: {original.Revisions.Count}");

        // Accept all revisions so the original becomes identical to the revised version.
        original.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException("All revisions should have been accepted.");
        }

        // Save the final document.
        string resultPath = Path.Combine(outputDir, "ComparedResult.docx");
        original.Save(resultPath);
        Console.WriteLine($"Result saved to: {resultPath}");
    }
}
