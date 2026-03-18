using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add paragraphs with identical formatting.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Writeln("Paragraph 1 - same format.");
        builder.Writeln("Paragraph 2 - same format.");

        // Change formatting for the next paragraph.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Writeln("Paragraph 3 - different format.");

        // Revert to the original formatting.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Writeln("Paragraph 4 - same as first format.");

        // Merge consecutive paragraphs that share identical formatting.
        MergeConsecutiveParagraphsWithSameFormatting(doc);

        // Save the resulting document.
        doc.Save("MergedParagraphs.docx");
    }

    // Merges consecutive paragraphs that have identical ParagraphFormat.
    static void MergeConsecutiveParagraphsWithSameFormatting(Document doc)
    {
        // Get the collection of paragraphs in the main story.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Iterate backwards so that removal does not affect the loop indices.
        for (int i = paragraphs.Count - 2; i >= 0; i--)
        {
            Paragraph current = paragraphs[i];
            Paragraph next = paragraphs[i + 1];

            if (ParagraphFormatsEqual(current.ParagraphFormat, next.ParagraphFormat))
            {
                // Move all child nodes (runs, fields, etc.) from the next paragraph
                // to the end of the current paragraph.
                while (next.HasChildNodes)
                {
                    Node child = next.FirstChild;
                    next.RemoveChild(child);
                    current.AppendChild(child);
                }

                // Remove the now empty paragraph.
                next.Remove();
            }
        }
    }

    // Determines whether two ParagraphFormat objects are equivalent for merging.
    static bool ParagraphFormatsEqual(ParagraphFormat a, ParagraphFormat b)
    {
        // Compare a subset of formatting properties that affect visual appearance.
        return a.StyleIdentifier == b.StyleIdentifier &&
               a.StyleName == b.StyleName &&
               a.Alignment == b.Alignment &&
               a.SpaceBefore == b.SpaceBefore &&
               a.SpaceAfter == b.SpaceAfter &&
               a.FirstLineIndent == b.FirstLineIndent &&
               a.LeftIndent == b.LeftIndent &&
               a.RightIndent == b.RightIndent &&
               a.LineSpacing == b.LineSpacing &&
               a.LineSpacingRule == b.LineSpacingRule &&
               a.KeepTogether == b.KeepTogether &&
               a.KeepWithNext == b.KeepWithNext;
    }
}
