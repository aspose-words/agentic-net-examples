using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Determine whether to enable tracking based on command‑line arguments.
        // Pass "on" as an argument to enable tracking, otherwise tracking stays disabled.
        bool enableTracking = args.Any(a => a.Equals("on", StringComparison.OrdinalIgnoreCase));

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content.
        builder.Writeln("Original paragraph.");

        if (enableTracking)
        {
            // Start tracking revisions.
            doc.StartTrackRevisions("DemoUser", DateTime.Now);

            // Perform some modifications that will be recorded as revisions.
            builder.Writeln("Added paragraph while tracking.");
            builder.MoveToDocumentStart();
            builder.Write("Inserted text at start. ");

            // Stop tracking revisions.
            doc.StopTrackRevisions();
        }

        // Save the document to a file.
        const string outputPath = "TrackedDocument.docx";
        doc.Save(outputPath);

        // Output revision information.
        Console.WriteLine($"Revisions count: {doc.Revisions.Count}");
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }
    }
}
