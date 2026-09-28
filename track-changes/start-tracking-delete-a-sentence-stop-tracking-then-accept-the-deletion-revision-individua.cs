using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build three separate paragraphs (sentences).
        builder.Writeln("This is the first sentence.");
        builder.Writeln("This is the sentence to delete.");
        builder.Writeln("This is the last sentence.");

        // Enable track changes.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Locate the paragraph that contains the sentence to delete.
        Paragraph paragraphToDelete = null;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.GetText().Contains("sentence to delete"))
            {
                paragraphToDelete = para;
                break;
            }
        }

        if (paragraphToDelete == null)
            throw new InvalidOperationException("Target paragraph not found.");

        // Remove the paragraph – this creates a deletion revision.
        paragraphToDelete.Remove();

        // Stop tracking changes.
        doc.StopTrackRevisions();

        // Accept only the deletion revision.
        // Iterate by index to avoid modifying the collection during enumeration.
        for (int i = doc.Revisions.Count - 1; i >= 0; i--)
        {
            Revision rev = doc.Revisions[i];
            if (rev.RevisionType == RevisionType.Deletion)
                rev.Accept();
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
