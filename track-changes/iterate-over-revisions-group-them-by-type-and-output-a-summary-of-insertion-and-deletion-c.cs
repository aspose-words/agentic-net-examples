using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add initial paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph 1.");
        builder.Writeln("Original paragraph 2.");

        // Enable track changes.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert a new paragraph (this will be recorded as an insertion revision).
        builder.Writeln("Inserted paragraph.");

        // Delete the first paragraph (this will be recorded as a deletion revision).
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        firstParagraph.Remove();

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document (optional, but ensures revisions are persisted).
        doc.Save("Revisions.docx");

        // Iterate over revisions and count insertions and deletions.
        int insertionCount = 0;
        int deletionCount = 0;

        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion)
                insertionCount++;
            else if (rev.RevisionType == RevisionType.Deletion)
                deletionCount++;
        }

        // Output the summary.
        Console.WriteLine($"Insertions: {insertionCount}");
        Console.WriteLine($"Deletions: {deletionCount}");
    }
}
