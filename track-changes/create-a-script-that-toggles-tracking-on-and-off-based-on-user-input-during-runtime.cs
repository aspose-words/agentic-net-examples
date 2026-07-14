using System;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not tracked.
        builder.Write("Initial content. ");

        // Determine whether to enable tracking based on the first command‑line argument.
        // Expected values: "on" to start tracking, anything else (or missing) leaves tracking disabled.
        bool enableTracking = args.Length > 0 && args[0].Equals("on", StringComparison.OrdinalIgnoreCase);

        if (enableTracking)
        {
            // Start tracking revisions with a specified author.
            doc.StartTrackRevisions("ConsoleUser", DateTime.Now);

            // This text will be recorded as an insertion revision.
            builder.Write("This text is tracked as a revision. ");

            // Stop tracking so subsequent changes are not recorded.
            doc.StopTrackRevisions();

            // This text will not be a revision.
            builder.Write("This text is added after tracking stopped. ");
        }
        else
        {
            // No tracking – all changes are normal edits.
            builder.Write("Tracking is disabled; this text is a normal edit. ");
        }

        // Inspect revisions if any were created.
        if (doc.HasRevisions)
        {
            Console.WriteLine($"Document has {doc.Revisions.Count} revision(s):");
            foreach (Revision rev in doc.Revisions)
            {
                Console.WriteLine($"- Author: {rev.Author}, Type: {rev.RevisionType}, Text: \"{rev.ParentNode.GetText().Trim()}\"");
            }

            // Accept all revisions to finalize the document.
            doc.AcceptAllRevisions();
        }
        else
        {
            Console.WriteLine("Document has no revisions.");
        }

        // Save the resulting document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
