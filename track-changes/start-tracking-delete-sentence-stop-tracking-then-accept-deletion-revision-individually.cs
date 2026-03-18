using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three paragraphs – the middle one will be deleted.
        builder.Writeln("First sentence.");
        builder.Writeln("Second sentence to be deleted.");
        builder.Writeln("Third sentence.");

        // Begin tracking revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Delete the second paragraph. The removal is recorded as a delete revision.
        Paragraph paragraphToDelete = doc.FirstSection.Body.Paragraphs[1];
        paragraphToDelete.Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Locate the delete revision and accept it individually.
        foreach (Revision revision in doc.Revisions)
        {
            if (revision.RevisionType == RevisionType.Deletion)
            {
                revision.Accept(); // Accept this specific deletion.
                break; // Only accept the first delete revision.
            }
        }

        // Save the final document.
        doc.Save("Output.docx");
    }
}
