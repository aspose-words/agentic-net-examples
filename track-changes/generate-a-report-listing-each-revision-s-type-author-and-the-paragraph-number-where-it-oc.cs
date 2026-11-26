using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace TrackChangesReport
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and add three paragraphs.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Paragraph 1");
            builder.Writeln("Paragraph 2");
            builder.Writeln("Paragraph 3");

            // Enable track changes.
            doc.StartTrackRevisions("Author1", DateTime.Now);

            // 1) Insert a new paragraph (creates an insertion revision).
            builder.Writeln("Inserted paragraph");

            // 2) Delete the second original paragraph (creates a deletion revision).
            Paragraph paraToDelete = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);
            paraToDelete.Remove();

            // 3) Change formatting of the third original paragraph (creates a format change revision).
            Paragraph paraToFormat = (Paragraph)doc.GetChild(NodeType.Paragraph, 2, true);
            // Apply bold formatting to the first run in the paragraph.
            Run run = (Run)paraToFormat.GetChild(NodeType.Run, 0, true);
            run.Font.Bold = true;

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Save the document (demonstrates that revisions are persisted).
            doc.Save("TrackedDocument.docx");

            // Generate report: list each revision's type, author, and paragraph number.
            RevisionCollection revisions = doc.Revisions;
            NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            for (int i = 0; i < revisions.Count; i++)
            {
                Revision rev = revisions[i];

                // Determine the paragraph that contains the revision.
                Paragraph containingParagraph = rev.ParentNode as Paragraph ??
                                                rev.ParentNode.GetAncestor(NodeType.Paragraph) as Paragraph;

                int paragraphNumber = containingParagraph != null
                    ? allParagraphs.IndexOf(containingParagraph) + 1 // 1‑based index
                    : -1; // Not inside a paragraph (should not happen for this example)

                Console.WriteLine($"Revision {i + 1}: Type={rev.RevisionType}, Author={rev.Author}, Paragraph={paragraphNumber}");
            }
        }
    }
}
