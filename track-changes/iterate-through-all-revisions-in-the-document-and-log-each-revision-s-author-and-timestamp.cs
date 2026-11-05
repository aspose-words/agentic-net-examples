using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content (these will be the original paragraphs).
        builder.Writeln("Original paragraph 1.");
        builder.Writeln("Original paragraph 2.");

        // Enable tracking of revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // 1. Insert a new paragraph – this will create an insertion revision.
        builder.Writeln("Inserted paragraph.");

        // 2. Delete the first original paragraph – this will create a deletion revision.
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        firstParagraph?.Remove();

        // 3. Change formatting of the remaining original paragraph – this will create a format change revision.
        Paragraph remainingParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        if (remainingParagraph != null)
        {
            Run run = (Run)remainingParagraph.GetChild(NodeType.Run, 0, true);
            if (run != null)
                run.Font.Bold = true;
        }

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document (required by the rules when output is expected).
        doc.Save("Revisions.docx");

        // Iterate through all revisions and log author and timestamp.
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"Revision Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }
    }
}
