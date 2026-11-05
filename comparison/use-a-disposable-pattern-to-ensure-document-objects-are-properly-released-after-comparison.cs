using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder originalBuilder = new DocumentBuilder(original);
        originalBuilder.Writeln("Hello world.");

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder revisedBuilder = new DocumentBuilder(revised);
        revisedBuilder.Writeln("Hello revised world.");

        // Perform comparison. Provide author name and current date/time.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Accept all revisions.
        original.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("All revisions should be accepted.");

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Comparison completed. Output saved to: {outputPath}");
    }
}
