using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content.
        builder.Writeln("Original paragraph.");

        // Enable track changes to generate revisions.
        doc.StartTrackRevisions("Sample Author", DateTime.Now);

        // Make a change that will be recorded as a revision.
        builder.Writeln("Inserted paragraph while tracking changes.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Check whether the document contains any revisions.
        bool hasRevisions = doc.HasRevisions;

        // Output the result.
        Console.WriteLine($"Document has revisions: {hasRevisions}");

        // (Optional) Save the document to verify revisions manually if needed.
        // doc.Save("RevisedDocument.docx");
    }
}
