using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable tracking of revisions.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert several paragraphs sequentially – these will be grouped as one revision group.
        builder.Writeln("First inserted paragraph.");
        builder.Writeln("Second inserted paragraph.");
        builder.Writeln("Third inserted paragraph.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Verify that a revision group was created.
        if (doc.Revisions.Groups.Count == 0)
            throw new InvalidOperationException("No revision groups were created.");

        // Accept all revisions in the document – this accepts the entire group in a single call.
        doc.Revisions.AcceptAll();

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
