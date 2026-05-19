using System;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Determine the mode based on command‑line arguments.
        // If no argument is supplied, default to "off".
        string mode = args.Length > 0 ? args[0].ToLowerInvariant() : "off";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not a revision.
        builder.Writeln("Initial content. ");

        if (mode == "on")
        {
            // Enable tracking of revisions.
            doc.StartTrackRevisions("User", DateTime.Now);

            // Add text while tracking is active – this will create an insertion revision.
            builder.Writeln("Tracked change: added while tracking is ON.");

            // Stop tracking so subsequent edits are not recorded as revisions.
            doc.StopTrackRevisions();

            // Save the document showing tracked changes.
            doc.Save("TrackingOn.docx");
        }
        else // mode == "off" or any other value
        {
            // Ensure tracking is not active.
            // (If it were previously started, we stop it.)
            doc.StopTrackRevisions();

            // Add text without tracking – no revisions will be created.
            builder.Writeln("Untracked change: added while tracking is OFF.");

            // Save the document without any revisions.
            doc.Save("TrackingOff.docx");
        }

        // For demonstration purposes, output the number of revisions created.
        Console.WriteLine($"Mode: {mode}");
        Console.WriteLine($"Revisions count: {doc.Revisions.Count}");
    }
}
