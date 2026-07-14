using System;
using Aspose.Words;

public class TrackChangesDemo
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three sentences – the middle one will be deleted while tracking.
        builder.Writeln("This is the first sentence.");
        builder.Writeln("This sentence will be deleted.");
        builder.Writeln("This is the last sentence.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Demo Author", DateTime.Now);

        // Delete the second paragraph (the sentence to be removed).
        // This operation creates a deletion-type revision.
        Paragraph paragraphToDelete = doc.FirstSection.Body.Paragraphs[1];
        paragraphToDelete.Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Find the deletion revision and accept it individually.
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Deletion)
            {
                rev.Accept(); // Accept only this deletion revision.
                break;
            }
        }

        // Save the resulting document.
        doc.Save("TrackChangesDemo.docx");

        // Output the final document text to the console (for verification).
        Console.WriteLine("Final document text:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
