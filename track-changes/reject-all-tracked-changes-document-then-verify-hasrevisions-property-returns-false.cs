using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Start tracking revisions and make some changes.
        doc.StartTrackRevisions("John Doe");
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");
        doc.StopTrackRevisions();

        // At this point the document has revisions.
        Console.WriteLine($"Has revisions before reject: {doc.HasRevisions}"); // Expected: True

        // Reject all revisions using the provided RevisionCollection.RejectAll method.
        doc.Revisions.RejectAll();

        // Verify that the document no longer has any revisions.
        Console.WriteLine($"Has revisions after reject: {doc.HasRevisions}"); // Expected: False
    }
}
