using System;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Determine whether to enable tracking based on the first command‑line argument.
        bool enableTracking = args.Length > 0 && args[0].Equals("on", StringComparison.OrdinalIgnoreCase);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content (not tracked).
        builder.Writeln("Initial content.");

        if (enableTracking)
        {
            // Start tracking revisions.
            doc.StartTrackRevisions("User", DateTime.Now);
            builder.Writeln("This text is added while tracking is ON.");
            // Stop tracking revisions.
            doc.StopTrackRevisions();
        }
        else
        {
            // Add content without tracking.
            builder.Writeln("This text is added while tracking is OFF.");
        }

        // Add final content after the toggle.
        builder.Writeln("Final content.");

        // Save the document to disk.
        doc.Save("ToggleTracking.docx");

        // Output the number of revisions created (non‑interactive).
        Console.WriteLine($"Revisions count: {doc.Revisions.Count}");
    }
}
