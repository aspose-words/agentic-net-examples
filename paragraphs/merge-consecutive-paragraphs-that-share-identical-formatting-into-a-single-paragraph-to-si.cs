using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample paragraphs with various formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph 1 – Bold, 12pt.
        builder.Font.Bold = true;
        builder.Font.Size = 12;
        builder.Writeln("First paragraph (Bold, 12pt).");

        // Paragraph 2 – Same formatting as paragraph 1 (should be merged).
        builder.Writeln("Second paragraph (Bold, 12pt).");

        // Paragraph 3 – Italic, 10pt (different).
        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Font.Size = 10;
        builder.Writeln("Third paragraph (Italic, 10pt).");

        // Paragraph 4 – Same formatting as paragraph 3 (should be merged).
        builder.Writeln("Fourth paragraph (Italic, 10pt).");

        // Paragraph 5 – Normal formatting (different).
        builder.Font.Italic = false;
        builder.Font.Size = 11;
        builder.Writeln("Fifth paragraph (Normal, 11pt).");

        // Save the original document (optional, for inspection).
        doc.Save("Original.docx");

        // Merge consecutive paragraphs that share identical formatting.
        MergeConsecutiveParagraphsWithSameFormatting(doc);

        // Save the merged document.
        doc.Save("Merged.docx");
    }

    private static void MergeConsecutiveParagraphsWithSameFormatting(Document doc)
    {
        // Get a live collection of all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Iterate backwards so that removal of a paragraph does not affect the index of yet‑to‑process items.
        for (int i = paragraphs.Count - 2; i >= 0; i--)
        {
            Paragraph current = (Paragraph)paragraphs[i];
            Paragraph next = (Paragraph)paragraphs[i + 1];

            if (ParagraphFormatsAreEqual(current, next))
            {
                // Append a space between the merged texts to keep words separated.
                current.AppendChild(new Run(doc, " "));

                // Move all child nodes (runs, fields, etc.) from the next paragraph into the current one.
                foreach (Node child in next.GetChildNodes(NodeType.Any, false))
                {
                    current.AppendChild(child.Clone(true));
                }

                // Remove the now‑empty next paragraph from the document.
                next.Remove();
            }
        }
    }

    private static bool ParagraphFormatsAreEqual(Paragraph p1, Paragraph p2)
    {
        ParagraphFormat f1 = p1.ParagraphFormat;
        ParagraphFormat f2 = p2.ParagraphFormat;

        // Compare the most common formatting properties.
        return
            f1.StyleIdentifier == f2.StyleIdentifier &&
            f1.Alignment == f2.Alignment &&
            f1.FirstLineIndent == f2.FirstLineIndent &&
            f1.LeftIndent == f2.LeftIndent &&
            f1.RightIndent == f2.RightIndent &&
            f1.LineSpacing == f2.LineSpacing &&
            f1.LineSpacingRule == f2.LineSpacingRule &&
            f1.SpaceAfter == f2.SpaceAfter &&
            f1.SpaceBefore == f2.SpaceBefore &&
            f1.KeepTogether == f2.KeepTogether &&
            f1.KeepWithNext == f2.KeepWithNext &&
            f1.WidowControl == f2.WidowControl &&
            // Compare the formatting of the first run (font) as a simple proxy for the whole paragraph.
            RunFormatsAreEqual(GetFirstRun(p1), GetFirstRun(p2));
    }

    private static Run GetFirstRun(Paragraph paragraph)
    {
        foreach (Node node in paragraph.GetChildNodes(NodeType.Run, false))
            return (Run)node;
        return null;
    }

    private static bool RunFormatsAreEqual(Run r1, Run r2)
    {
        // If both paragraphs have no runs, consider them equal.
        if (r1 == null && r2 == null) return true;
        if (r1 == null || r2 == null) return false;

        Font f1 = r1.Font;
        Font f2 = r2.Font;

        return
            f1.Name == f2.Name &&
            f1.Size == f2.Size &&
            f1.Bold == f2.Bold &&
            f1.Italic == f2.Italic &&
            f1.Underline == f2.Underline &&
            f1.Color.Equals(f2.Color);
    }
}
