using System;
using System.IO;
using Aspose.Words;

public class ComparisonAcceptRevisionsExample
{
    public static void Main()
    {
        // Create the original document with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");

        // Create the revised document with a modification.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised paragraph with changes.");

        // Compare the two documents. The comparison will generate revisions in the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Accept all revisions, cleaning the document.
        original.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException("All revisions should have been accepted.");
        }

        // Save the cleaned document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "cleaned.docx");
        original.Save(outputPath);
    }
}
