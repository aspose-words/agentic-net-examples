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
        builder.Writeln("This is the first sentence.");
        builder.Writeln("This is the sentence to delete.");
        builder.Writeln("This is the last sentence.");

        // Enable track changes.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Delete the second paragraph while tracking is enabled.
        Paragraph paragraphToDelete = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[1];
        paragraphToDelete.Remove();

        // Stop tracking changes.
        doc.StopTrackRevisions();

        // Verify that a deletion revision exists and accept it individually.
        bool deletionFound = false;
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Deletion)
            {
                rev.Accept(); // Accept this specific deletion revision.
                deletionFound = true;
                break;
            }
        }

        if (!deletionFound)
            throw new InvalidOperationException("Expected deletion revision was not found.");

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
