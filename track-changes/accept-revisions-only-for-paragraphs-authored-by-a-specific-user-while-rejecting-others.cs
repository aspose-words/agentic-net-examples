using System;
using Aspose.Words; // Revision, Document, DocumentBuilder are in this namespace

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that is not a revision.
        builder.Writeln("This is the original paragraph.");

        // Track changes made by Alice.
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("Paragraph added by Alice.");
        doc.StopTrackRevisions();

        // Track changes made by Bob.
        doc.StartTrackRevisions("Bob", DateTime.Now);
        builder.Writeln("Paragraph added by Bob.");
        doc.StopTrackRevisions();

        // Ensure that revisions were created.
        if (!doc.HasRevisions)
            throw new InvalidOperationException("No revisions were generated.");

        // Accept revisions authored by Alice, reject all others.
        // Iterate backwards because accepting/rejecting modifies the collection.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.Author == "Alice")
                rev.Accept();
            else
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
