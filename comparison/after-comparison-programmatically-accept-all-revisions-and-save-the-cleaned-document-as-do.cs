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
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Compare the documents – revisions will be added to the original document.
        original.Compare(revised, "Author", DateTime.Now);

        // Ensure that at least one revision was created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison.");

        // Accept all revisions, turning the original into the revised version.
        original.AcceptAllRevisions();

        // Verify that all revisions have been cleared.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were not fully accepted.");

        // Save the cleaned document.
        original.Save("cleaned.docx");
    }
}
