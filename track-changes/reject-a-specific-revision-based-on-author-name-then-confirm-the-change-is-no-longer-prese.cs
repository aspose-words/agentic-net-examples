using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that is not a revision.
        builder.Writeln("Original paragraph.");

        // Track changes made by John Doe.
        doc.StartTrackRevisions("John Doe", DateTime.Now);
        builder.Writeln("Paragraph added by John.");
        doc.StopTrackRevisions();

        // Track changes made by Jane Smith.
        doc.StartTrackRevisions("Jane Smith", DateTime.Now);
        builder.Writeln("Paragraph added by Jane.");
        doc.StopTrackRevisions();

        // Collect revisions authored by John Doe.
        List<Revision> revisionsToReject = new List<Revision>();
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.Author == "John Doe")
                revisionsToReject.Add(rev);
        }

        // Reject the collected revisions.
        foreach (Revision rev in revisionsToReject)
        {
            rev.Reject();
        }

        // Verify that no revision from John Doe remains.
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.Author == "John Doe")
                throw new InvalidOperationException("John Doe revision was not removed.");
        }

        // Save the document to demonstrate the result.
        doc.Save("Result.docx");

        // Output confirmation.
        Console.WriteLine($"Revisions by John Doe have been rejected. Remaining revisions: {doc.Revisions.Count}");
    }
}
