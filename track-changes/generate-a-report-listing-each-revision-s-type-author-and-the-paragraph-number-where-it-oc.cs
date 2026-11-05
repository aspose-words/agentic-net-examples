using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial paragraphs.
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // 1. Insert a new paragraph after the first one.
        Paragraph firstPara = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        builder.MoveTo(firstPara);
        builder.Writeln("Inserted Paragraph");

        // 2. Delete the original second paragraph.
        Paragraph secondPara = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);
        secondPara.Remove();

        // 3. Change formatting of the third paragraph (make it bold).
        // The third paragraph is now at index 2 because we removed one paragraph.
        Paragraph thirdPara = (Paragraph)doc.GetChild(NodeType.Paragraph, 2, true);
        foreach (Run run in thirdPara.Runs)
        {
            run.Font.Bold = true; // This creates a format‑change revision.
        }

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Generate report of revisions.
        Console.WriteLine("Revision Report:");
        foreach (Revision rev in doc.Revisions)
        {
            // Find the paragraph that contains the revision.
            Node node = rev.ParentNode;
            while (node != null && !(node is Paragraph))
                node = node.ParentNode;

            int paragraphNumber = -1;
            if (node != null)
            {
                NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                paragraphNumber = allParagraphs.IndexOf(node) + 1; // 1‑based index
            }

            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Paragraph: {paragraphNumber}");
        }

        // Save the document to verify revisions visually.
        doc.Save("TrackedDocument.docx");
    }
}
