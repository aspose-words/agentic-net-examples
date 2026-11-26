using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph.");

        // Determine whether tracking should be enabled based on the first command‑line argument.
        bool enableTracking = args.Length > 0 && args[0].Equals("on", StringComparison.OrdinalIgnoreCase);

        if (enableTracking)
        {
            // Start tracking revisions.
            doc.StartTrackRevisions("DemoUser", DateTime.Now);

            // Make a change that will be recorded as a revision.
            builder.Writeln("Paragraph added while tracking is enabled.");

            // Stop tracking revisions.
            doc.StopTrackRevisions();
        }
        else
        {
            // Tracking is not enabled; make a change that will NOT be recorded.
            builder.Writeln("Paragraph added while tracking is disabled.");
        }

        // Inspect revisions.
        RevisionCollection revisions = doc.Revisions;
        Console.WriteLine($"Tracking enabled: {enableTracking}");
        Console.WriteLine($"Number of revisions in the document: {revisions.Count}");

        // Output details of each revision (if any).
        foreach (Revision rev in revisions)
        {
            Console.WriteLine($"Revision Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }

        // Save the document to the output file.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
