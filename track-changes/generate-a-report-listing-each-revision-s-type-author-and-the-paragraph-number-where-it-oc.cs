using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add initial paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Enable track changes.
        doc.StartTrackRevisions("AuthorName", DateTime.Now);

        // Insert a new paragraph (this will be an insertion revision).
        builder.Writeln("Inserted paragraph while tracking.");

        // Modify text of the first paragraph (creates insertion/deletion revisions).
        Paragraph firstPara = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        if (firstPara.Runs.Count > 0)
            firstPara.Runs[0].Text = "First paragraph modified.";

        // Change formatting of the second paragraph (format change revision).
        Paragraph secondPara = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);
        secondPara.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Stop tracking changes.
        doc.StopTrackRevisions();

        // Save the document (required by lifecycle rule).
        doc.Save("SampleRevisions.docx");

        // Prepare a collection of all paragraphs for paragraph number lookup.
        NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Generate report of revisions.
        int revIndex = 1;
        foreach (Revision rev in doc.Revisions)
        {
            // Determine the paragraph that contains the revision.
            Paragraph para = rev.ParentNode as Paragraph ?? rev.ParentNode?.GetAncestor(NodeType.Paragraph) as Paragraph;
            int paragraphNumber = -1;
            if (para != null)
                paragraphNumber = allParagraphs.IndexOf(para) + 1; // 1‑based index

            string paraInfo = paragraphNumber > 0 ? paragraphNumber.ToString() : "N/A";

            Console.WriteLine($"Revision {revIndex}: Type={rev.RevisionType}, Author={rev.Author}, Paragraph={paraInfo}");
            revIndex++;
        }
    }
}
