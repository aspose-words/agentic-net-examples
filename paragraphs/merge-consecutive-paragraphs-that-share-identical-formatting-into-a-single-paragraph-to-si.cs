using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add sample paragraphs with various formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph 1 – default formatting.
        builder.Writeln("Paragraph 1 – default formatting.");

        // Paragraph 2 – bold formatting.
        builder.Font.Bold = true;
        builder.Writeln("Paragraph 2 – bold formatting.");

        // Paragraph 3 – same bold formatting as previous paragraph.
        builder.Writeln("Paragraph 3 – also bold.");

        // Paragraph 4 – back to default formatting.
        builder.Font.Bold = false;
        builder.Writeln("Paragraph 4 – default formatting again.");

        // Paragraph 5 – italic formatting.
        builder.Font.Italic = true;
        builder.Writeln("Paragraph 5 – italic.");

        // Paragraph 6 – same italic formatting as previous paragraph.
        builder.Writeln("Paragraph 6 – also italic.");

        // Reset formatting for any further content.
        builder.Font.Italic = false;

        // Merge consecutive paragraphs that share identical formatting.
        MergeConsecutiveParagraphsWithSameFormatting(doc);

        // Save the resulting document.
        doc.Save("MergedParagraphs.docx");
    }

    private static void MergeConsecutiveParagraphsWithSameFormatting(Document doc)
    {
        // Work from the end towards the start so that removal of paragraphs does not affect the index of yet‑to‑process items.
        for (int i = doc.FirstSection.Body.Paragraphs.Count - 1; i > 0; i--)
        {
            Paragraph current = doc.FirstSection.Body.Paragraphs[i];
            Paragraph previous = doc.FirstSection.Body.Paragraphs[i - 1];

            if (HaveIdenticalFormatting(previous, current))
            {
                // Move all child nodes (runs, fields, etc.) from the current paragraph to the previous one.
                while (current.HasChildNodes)
                {
                    Node child = current.FirstChild;
                    current.RemoveChild(child);
                    previous.AppendChild(child);
                }

                // Remove the now‑empty current paragraph.
                current.Remove();
            }
        }
    }

    private static bool HaveIdenticalFormatting(Paragraph p1, Paragraph p2)
    {
        ParagraphFormat f1 = p1.ParagraphFormat;
        ParagraphFormat f2 = p2.ParagraphFormat;

        // Compare style identifiers and names.
        if (f1.StyleIdentifier != f2.StyleIdentifier) return false;
        if (f1.StyleName != f2.StyleName) return false;

        // Compare common formatting properties.
        if (f1.Alignment != f2.Alignment) return false;
        if (f1.LeftIndent != f2.LeftIndent) return false;
        if (f1.RightIndent != f2.RightIndent) return false;
        if (f1.SpaceAfter != f2.SpaceAfter) return false;
        if (f1.SpaceBefore != f2.SpaceBefore) return false;
        if (f1.LineSpacing != f2.LineSpacing) return false;
        if (f1.LineSpacingRule != f2.LineSpacingRule) return false;

        // If all compared properties are equal, consider the formatting identical.
        return true;
    }
}
