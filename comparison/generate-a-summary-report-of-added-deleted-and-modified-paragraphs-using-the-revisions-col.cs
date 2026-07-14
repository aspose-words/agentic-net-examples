using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with three paragraphs.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Paragraph 1.");
        builder.Writeln("Paragraph 2.");
        builder.Writeln("Paragraph 3.");

        // Create the revised document with modifications:
        // - Paragraph 2 is changed.
        // - Paragraph 3 is removed.
        // - Paragraph 4 is added.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Paragraph 1.");               // unchanged
        builder2.Writeln("Paragraph 2 modified.");      // modified
        builder2.Writeln("Paragraph 4.");               // added

        // Compare the documents. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Summarize revisions.
        int addedParagraphs = 0;
        int deletedParagraphs = 0;
        HashSet<string> modifiedParagraphs = new HashSet<string>();

        foreach (Revision rev in original.Revisions)
        {
            Node parent = rev.ParentNode;
            if (parent == null) continue;

            // Count insertions and deletions that occur at the paragraph level.
            if (parent.NodeType == NodeType.Paragraph)
            {
                if (rev.RevisionType == RevisionType.Insertion)
                    addedParagraphs++;
                else if (rev.RevisionType == RevisionType.Deletion)
                    deletedParagraphs++;
            }
            else
            {
                // For run‑level changes, treat the containing paragraph as modified.
                Paragraph para = (Paragraph)parent.GetAncestor(NodeType.Paragraph);
                if (para != null)
                    modifiedParagraphs.Add(para.GetText().Trim());
            }
        }

        // Build the summary report.
        string report = $"Added paragraphs: {addedParagraphs}{Environment.NewLine}" +
                        $"Deleted paragraphs: {deletedParagraphs}{Environment.NewLine}" +
                        $"Modified paragraphs: {modifiedParagraphs.Count}{Environment.NewLine}";

        foreach (string txt in modifiedParagraphs)
            report += $" - {txt}{Environment.NewLine}";

        // Output the report to the console.
        Console.WriteLine(report);

        // Save the report to a text file.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisionSummary.txt");
        File.WriteAllText(reportPath, report);

        // Save the compared document for reference.
        string comparedPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        original.Save(comparedPath);
    }
}
