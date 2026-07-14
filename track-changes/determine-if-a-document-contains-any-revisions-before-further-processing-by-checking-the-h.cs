using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to edit it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text without tracking – this will not be a revision.
        builder.Writeln("Original text.");

        // Enable tracking of revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Write text while tracking is enabled – this will be recorded as a revision.
        builder.Writeln("Added revision text.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Determine whether the document contains any revisions.
        bool hasRevisions = doc.HasRevisions;

        // Output the result.
        Console.WriteLine($"Document has revisions: {hasRevisions}");

        // Save the document (optional, demonstrates lifecycle usage).
        doc.Save("RevisionsCheck.docx");
    }
}
