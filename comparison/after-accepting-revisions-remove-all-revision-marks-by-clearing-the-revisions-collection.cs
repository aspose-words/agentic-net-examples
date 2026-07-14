using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original content.");

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised content.");

        // Compare the documents – revisions will be added to the original document.
        original.Compare(revised, "Author", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison.");

        // Accept all revisions.
        original.Revisions.AcceptAll();

        // After accepting, the revisions collection should be empty.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("All revisions should have been cleared.");

        // Save the final document without any revision marks.
        original.Save("Result.docx");
    }
}
