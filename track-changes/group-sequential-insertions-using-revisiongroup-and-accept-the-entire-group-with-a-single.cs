using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content.
        builder.Writeln("Hello");

        // Start tracking revisions.
        doc.StartTrackRevisions("Author1", DateTime.Now);

        // Insert sequential text which will be treated as separate insertions.
        builder.Write(" World");
        builder.Write("!");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // At this point the document should contain revisions for the two insertions.
        RevisionCollection revisions = doc.Revisions;
        Console.WriteLine($"Total revisions before acceptance: {revisions.Count}");

        // Accept all revisions in one call (Aspose.Words groups sequential insertions internally).
        doc.AcceptAllRevisions();

        // Verify that revisions have been accepted.
        Console.WriteLine($"Total revisions after acceptance: {doc.Revisions.Count}");

        // Save the resulting document.
        doc.Save("Result.docx");
        Console.WriteLine("Document saved as Result.docx");
    }
}
