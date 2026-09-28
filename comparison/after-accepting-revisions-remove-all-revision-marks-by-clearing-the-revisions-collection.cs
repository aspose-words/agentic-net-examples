using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original text.");

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised text with a change.");

        // Compare the documents to generate revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison.");

        // Accept all revisions, which removes revision marks and clears the collection.
        original.AcceptAllRevisions();

        // Verify that the revisions collection is now empty.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions collection should be empty after accepting all revisions.");

        // Save the final document.
        original.Save("final.docx");
    }
}
