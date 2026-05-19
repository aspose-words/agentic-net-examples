using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not a revision.
        builder.Writeln("Original paragraph.");

        // First author: Alice – start tracking revisions.
        string targetAuthor = "Alice";
        doc.StartTrackRevisions(targetAuthor, DateTime.Now);
        builder.Writeln("Paragraph added by Alice."); // Revision authored by Alice.
        doc.StopTrackRevisions();

        // Second author: Bob – start tracking revisions.
        string otherAuthor = "Bob";
        doc.StartTrackRevisions(otherAuthor, DateTime.Now);
        builder.Writeln("Paragraph added by Bob."); // Revision authored by Bob.
        doc.StopTrackRevisions();

        // Accept only revisions authored by Alice and reject all others.
        // Iterate backwards to avoid modifying the collection while enumerating.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.Author == targetAuthor)
                rev.Accept();
            else
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
