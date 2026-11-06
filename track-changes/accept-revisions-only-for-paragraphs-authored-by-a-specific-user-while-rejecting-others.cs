using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an initial paragraph (no revisions).
        builder.Writeln("Original paragraph.");

        // ---------- Revisions by Alice ----------
        // Start tracking revisions with author "Alice".
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Insert a new paragraph authored by Alice.
        builder.Writeln("Paragraph added by Alice.");

        // Modify the original paragraph (insertion) authored by Alice.
        builder.MoveToDocumentStart();
        builder.Write("Alice inserted text. ");

        // Stop tracking for Alice.
        doc.StopTrackRevisions();

        // ---------- Revisions by Bob ----------
        // Start tracking revisions with author "Bob".
        doc.StartTrackRevisions("Bob", DateTime.Now);

        // Insert a new paragraph authored by Bob.
        builder.Writeln("Paragraph added by Bob.");

        // Delete the original paragraph (creates a deletion revision by Bob).
        Paragraph originalPara = doc.FirstSection.Body.Paragraphs[0];
        originalPara.Remove();

        // Stop tracking for Bob.
        doc.StopTrackRevisions();

        // ---------- Selective acceptance/rejection ----------
        // Accept only revisions authored by "Alice", reject all others.
        // Create a snapshot of the revisions to avoid modifying the collection while iterating.
        var revisions = doc.Revisions.Cast<Revision>().ToList();

        foreach (Revision rev in revisions)
        {
            if (rev.Author == "Alice")
                rev.Accept();
            else
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
