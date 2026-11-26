using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an initial paragraph (not a revision).
        builder.Writeln("Initial paragraph.");

        // First author makes a revision.
        doc.StartTrackRevisions("Alice", DateTime.Now);
        builder.Writeln("Paragraph added by Alice.");
        doc.StopTrackRevisions();

        // Second author makes a revision.
        doc.StartTrackRevisions("Bob", DateTime.Now);
        builder.Writeln("Paragraph added by Bob.");
        doc.StopTrackRevisions();

        // Ensure revisions were created.
        if (doc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated.");

        // Accept revisions authored by Alice, reject all others.
        Revision[] revisions = doc.Revisions.ToArray();
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
