using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three separate paragraphs (sentences).
        builder.Writeln("First sentence.");
        builder.Writeln("Second sentence.");
        builder.Writeln("Third sentence.");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Delete the second paragraph while tracking is enabled.
        // This will create a deletion-type revision.
        Paragraph secondParagraph = doc.FirstSection.Body.Paragraphs[1];
        secondParagraph.Remove();

        // Stop tracking revisions.
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

        // Indicate completion (optional).
        Console.WriteLine("Document processed and saved.");
    }
}
