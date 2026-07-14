using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original document.");

        // Create the edited document with a different paragraph.
        Document edited = new Document();
        builder = new DocumentBuilder(edited);
        builder.Writeln("This is the edited document.");

        // Ensure both documents have no revisions before comparison.
        if (original.Revisions.Count != 0 || edited.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. The original document will receive revisions that represent the differences.
        original.Compare(edited, "John Doe", DateTime.Now);

        // Save the original document that now contains revision markup.
        string revisionsPath = Path.Combine(Directory.GetCurrentDirectory(), "OriginalWithRevisions.docx");
        original.Save(revisionsPath);

        // Output revision details to the console.
        Console.WriteLine("Revisions detected:");
        foreach (Revision rev in original.Revisions)
        {
            Console.WriteLine($"- Type: {rev.RevisionType}, Author: {rev.Author}, Text: \"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Accept all revisions so the original document becomes identical to the edited one.
        original.Revisions.AcceptAll();

        // Save the accepted version.
        string acceptedPath = Path.Combine(Directory.GetCurrentDirectory(), "Accepted.docx");
        original.Save(acceptedPath);
    }
}
