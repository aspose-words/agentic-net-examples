using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text that will not be tracked as a revision.
        builder.Writeln("Paragraph before tracking.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Insert a new paragraph – this will be an insertion revision.
        builder.Writeln("This paragraph is an insertion revision.");

        // Delete the first paragraph – this will create a deletion revision.
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        firstParagraph.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Reject only deletion revisions, keep insertions and formatting changes.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.RevisionType == RevisionType.Deletion)
                rev.Reject();
        }

        // Save the resulting document.
        doc.Save("Output.docx");

        // Optional: display the final document text.
        Console.WriteLine("Final document text:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
