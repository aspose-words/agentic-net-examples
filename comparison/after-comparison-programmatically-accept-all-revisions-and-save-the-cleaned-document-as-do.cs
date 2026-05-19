using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original paragraph.");

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised paragraph with a change.");

        // Compare the documents – revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision was created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison.");

        // Accept all revisions, turning the original into the revised version.
        original.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("All revisions should have been accepted.");

        // Save the cleaned document.
        string outputPath = "Cleaned.docx";
        original.Save(outputPath);
    }
}
