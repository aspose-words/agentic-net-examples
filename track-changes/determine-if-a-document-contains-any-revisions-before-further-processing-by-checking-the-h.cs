using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph.");

        // Enable track changes.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Make a change that will be recorded as a revision.
        builder.Writeln("Added paragraph while tracking changes.");

        // Stop tracking.
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates persistence).
        string filePath = "RevisionsDemo.docx";
        doc.Save(filePath);

        // Determine if the document contains any revisions.
        bool hasRevisions = doc.HasRevisions;

        // Output the result.
        Console.WriteLine($"Document has revisions: {hasRevisions}");
    }
}
