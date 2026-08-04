using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the first document with deterministic content.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("This is a sample paragraph for comparison.");

        // Create the second document with exactly the same content.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("This is a sample paragraph for comparison.");

        // Ensure both documents have no revisions before comparison.
        if (doc1.Revisions.Count != 0 || doc2.Revisions.Count != 0)
            throw new InvalidOperationException("Documents should not contain revisions before comparison.");

        // Compare the two identical documents.
        doc1.Compare(doc2, "Author", DateTime.Now);

        // Verify that the comparison produced zero revisions, meaning the documents are identical.
        if (doc1.Revisions.Count != 0)
            throw new InvalidOperationException("Expected zero revisions after comparing identical documents.");

        // Save the result (optional, demonstrates that the API works end‑to‑end).
        doc1.Save("identical_compare.docx");

        // Inform the user that the documents are identical.
        Console.WriteLine("Documents are identical; no revisions were created.");
    }
}
