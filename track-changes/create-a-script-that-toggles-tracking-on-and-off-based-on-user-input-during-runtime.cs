using System;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Determine initial tracking state from the first argument (true/false). Default is false.
        bool startTracking = false;
        if (args.Length > 0 && bool.TryParse(args[0], out bool parsed))
            startTracking = parsed;

        // Create a new document with an initial paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph.");

        // ---------- First phase ----------
        // Enable tracking if requested.
        if (startTracking)
            doc.StartTrackRevisions("User", DateTime.Now);

        // Make a change that will be recorded as a revision when tracking is on.
        builder.Writeln("First change - may be a revision.");

        // Stop tracking if it was started.
        if (startTracking)
            doc.StopTrackRevisions();

        // ---------- Second phase ----------
        // Toggle tracking to the opposite state.
        if (!startTracking)
            doc.StartTrackRevisions("User", DateTime.Now);

        // Make another change.
        builder.Writeln("Second change - may be a revision.");

        // Ensure tracking is stopped at the end of the document modifications.
        if (!startTracking)
            doc.StopTrackRevisions();

        // Save the document containing the revisions.
        string filePath = "TrackedDocument.docx";
        doc.Save(filePath);

        // Inspect and output revision information.
        Console.WriteLine($"Total revisions in document: {doc.Revisions.Count}");
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }

        // Accept all revisions and save the final version.
        doc.AcceptAllRevisions();
        doc.Save("TrackedDocument_Accepted.docx");
    }
}
