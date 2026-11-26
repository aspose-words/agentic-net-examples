using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original content.");

        // Enable track changes to generate revisions.
        doc.StartTrackRevisions("Sample Author", DateTime.Now);

        // Perform a modification that will be recorded as a revision.
        builder.Writeln("Added line while tracking changes.");

        // Stop tracking changes.
        doc.StopTrackRevisions();

        // Save the document (optional, demonstrates file output).
        string filePath = "TrackedDocument.docx";
        doc.Save(filePath);

        // Determine if the document contains any revisions.
        bool hasRevisions = doc.HasRevisions;

        // Output the result.
        Console.WriteLine(hasRevisions ? "Document has revisions." : "Document has no revisions.");
    }
}
