using System;
using Aspose.Words;

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

        // Compare the documents – revisions will be added to the original.
        original.Compare(revised, "Author", DateTime.Now);

        // Ensure that revisions were created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison.");

        // Accept all revisions, which removes the revision marks.
        original.Revisions.AcceptAll();

        // Verify that the revisions collection is now empty.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were not cleared after acceptance.");

        // Save the resulting document.
        original.Save("final.docx");
    }
}
