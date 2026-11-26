using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content (no revisions).
        builder.Writeln("Initial content.");

        // ---------- Create an old revision ----------
        // Start tracking with an old date.
        DateTime oldRevisionDate = DateTime.Now.AddDays(-10);
        doc.StartTrackRevisions("OldAuthor", oldRevisionDate);
        // Insert a paragraph that will be recorded as an insertion revision.
        builder.Writeln("This is an old revision that should be rejected.");
        // Stop tracking.
        doc.StopTrackRevisions();

        // ---------- Create a newer revision ----------
        // Start tracking with a newer date.
        DateTime newRevisionDate = DateTime.Now.AddDays(-1);
        doc.StartTrackRevisions("NewAuthor", newRevisionDate);
        // Insert a paragraph that will be recorded as an insertion revision.
        builder.Writeln("This is a new revision that should be kept.");
        // Stop tracking.
        doc.StopTrackRevisions();

        // Define the cutoff date: revisions older than this will be rejected.
        DateTime cutoffDate = DateTime.Now.AddDays(-5);

        // Collect revisions to a list to avoid modifying the collection while iterating.
        List<Revision> revisions = new List<Revision>(doc.Revisions);

        int rejectedCount = 0;
        int keptCount = 0;

        foreach (Revision rev in revisions)
        {
            if (rev.DateTime < cutoffDate)
            {
                rev.Reject(); // Reject old revision.
                rejectedCount++;
            }
            else
            {
                // Keep newer revision (do nothing).
                keptCount++;
            }
        }

        // Save the resulting document.
        doc.Save("Output.docx");

        // Output simple statistics (optional, not required for functionality).
        Console.WriteLine($"Total revisions processed: {revisions.Count}");
        Console.WriteLine($"Revisions rejected (older than {cutoffDate:d}): {rejectedCount}");
        Console.WriteLine($"Revisions kept: {keptCount}");
    }
}
