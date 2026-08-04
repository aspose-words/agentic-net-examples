using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text without tracking.
        builder.Writeln("Paragraph before tracking.");

        // Simulated user input: first toggle = true (enable tracking).
        bool enableTracking = true;
        if (enableTracking)
        {
            // Start tracking revisions with a specific author.
            doc.StartTrackRevisions("Alice", DateTime.Now);
        }

        // Add text while tracking is enabled – this will create a revision.
        builder.Writeln("Paragraph added while tracking is ON.");

        // Simulated user input: second toggle = false (disable tracking).
        bool disableTracking = true;
        if (disableTracking)
        {
            // Stop tracking revisions.
            doc.StopTrackRevisions();
        }

        // Add more text after tracking is stopped – this will NOT create a revision.
        builder.Writeln("Paragraph added after tracking is OFF.");

        // Inspect the revisions collection.
        int revisionCount = doc.Revisions.Count;
        Console.WriteLine($"Total revisions in the document: {revisionCount}");

        // Output details of each revision.
        for (int i = 0; i < revisionCount; i++)
        {
            Revision rev = doc.Revisions[i];
            Console.WriteLine($"Revision {i + 1}: Type={rev.RevisionType}, Author={rev.Author}, Text=\"{rev.ParentNode.GetText().Trim()}\"");
        }

        // Save the document to verify the changes.
        doc.Save("TrackChangesDemo.docx");
    }
}
