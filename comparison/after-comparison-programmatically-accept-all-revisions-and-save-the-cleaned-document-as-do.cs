using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with some text.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");
        builder1.Writeln("This line will stay the same.");

        // Create the revised document that contains a change.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello changed world."); // Modified line.
        builder2.Writeln("This line will stay the same."); // Unchanged line.

        // Compare the documents. Revisions are added to the original document.
        original.Compare(revised, "Author", DateTime.Now);

        // Ensure that the comparison produced at least one revision.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison.");

        // Accept all revisions, turning the original into the revised version.
        original.Revisions.AcceptAll();

        // Verify that all revisions have been accepted.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("All revisions should be accepted.");

        // Save the cleaned (revision‑free) document.
        original.Save("CleanedDocument.docx");
    }
}
