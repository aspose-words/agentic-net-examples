using System;
using Aspose.Words;

public class RevisionRemovalExample
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Compare the documents to generate revisions.
        original.Compare(revised, "ExampleAuthor", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Accept all revisions.
        original.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (original.Revisions.Count != 0)
        {
            throw new InvalidOperationException("All revisions should be accepted.");
        }

        // Save the final document without any revision marks.
        string outputPath = "final.docx";
        original.Save(outputPath);
    }
}
