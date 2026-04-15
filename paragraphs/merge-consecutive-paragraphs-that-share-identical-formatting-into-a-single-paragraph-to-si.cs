using System;
using Aspose.Words;

namespace MergeParagraphsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert paragraphs with identical formatting (Heading1) and one with different formatting (Normal).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("First paragraph – Heading 1");
            builder.Writeln("Second paragraph – Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Third paragraph – Normal style");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Fourth paragraph – Heading 1");

            // Merge consecutive paragraphs that share the same formatting.
            MergeConsecutiveParagraphs(doc);

            // Optional: join runs with the same formatting inside the merged paragraphs.
            doc.JoinRunsWithSameFormatting();

            // Save the resulting document.
            doc.Save("MergedParagraphs.docx");
        }

        private static void MergeConsecutiveParagraphs(Document doc)
        {
            // Get the collection of paragraphs in the first section's body.
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            // Iterate through the collection, merging when formatting matches.
            for (int i = 0; i < paragraphs.Count - 1; )
            {
                Paragraph current = paragraphs[i];
                Paragraph next = paragraphs[i + 1];

                if (HaveSameFormatting(current, next))
                {
                    // Move all child nodes (runs, fields, etc.) from the next paragraph into the current one.
                    while (next.HasChildNodes)
                    {
                        Node child = next.FirstChild;
                        next.RemoveChild(child);
                        current.AppendChild(child);
                    }

                    // Remove the now-empty next paragraph from the document.
                    next.Remove();

                    // Do not increment i – the current paragraph may still match the new next paragraph.
                }
                else
                {
                    // Formatting differs; move to the next pair.
                    i++;
                }
            }
        }

        private static bool HaveSameFormatting(Paragraph p1, Paragraph p2)
        {
            ParagraphFormat f1 = p1.ParagraphFormat;
            ParagraphFormat f2 = p2.ParagraphFormat;

            // Compare a subset of formatting properties that define the paragraph's appearance.
            return f1.StyleIdentifier == f2.StyleIdentifier &&
                   f1.Alignment == f2.Alignment &&
                   f1.FirstLineIndent == f2.FirstLineIndent &&
                   f1.LeftIndent == f2.LeftIndent &&
                   f1.RightIndent == f2.RightIndent;
        }
    }
}
