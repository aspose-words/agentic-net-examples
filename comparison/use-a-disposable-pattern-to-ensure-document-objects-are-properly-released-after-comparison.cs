using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");

        // Create the revised document with a deliberate difference.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello revised world.");

        // Compare the two documents. The comparison creates revision marks in 'original'.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision was generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Accept all revisions to clean the document.
        original.AcceptAllRevisions();

        // Ensure that no revisions remain after acceptance.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException("All revisions should be accepted.");
        }

        // Save the final document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);
    }
}
