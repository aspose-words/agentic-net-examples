using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph 1 – Normal style, left aligned.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Writeln("First paragraph with normal style.");

        // Paragraph 2 – Same formatting as paragraph 1 (should be merged).
        builder.Writeln("Second paragraph shares the same formatting.");

        // Paragraph 3 – Different style (will stay separate).
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Writeln("A heading paragraph with different formatting.");

        // Paragraph 4 – Same formatting as paragraph 3 (should be merged with it).
        builder.Writeln("Another heading paragraph with the same formatting.");

        // Paragraph 5 – Back to Normal style, left aligned (should merge with paragraph 1/2).
        builder.ParagraphFormat.StyleName = "Normal";
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Writeln("Third normal paragraph, should merge with the first group.");

        // Perform merging of consecutive paragraphs that share identical formatting.
        MergeConsecutiveParagraphs(doc);

        // Save the resulting document.
        doc.Save("MergedParagraphs.docx");
    }

    private static void MergeConsecutiveParagraphs(Document doc)
    {
        // Get the collection of paragraphs in the main body of the first section.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        int i = 0;
        while (i < paragraphs.Count - 1)
        {
            Paragraph current = paragraphs[i];
            Paragraph next = paragraphs[i + 1];

            if (HaveSameFormatting(current, next))
            {
                // Insert a space between the merged texts to keep readability.
                Run spaceRun = new Run(doc, " ");
                current.AppendChild(spaceRun);

                // Move all runs from the next paragraph into the current one.
                foreach (Run run in next.Runs)
                {
                    // Clone the run to preserve its formatting and add it to the current paragraph.
                    current.AppendChild(run.Clone(true));
                }

                // Remove the now-empty next paragraph.
                next.Remove();

                // Do not increment i to check if further consecutive paragraphs also match.
            }
            else
            {
                i++; // Move to the next paragraph pair.
            }
        }
    }

    private static bool HaveSameFormatting(Paragraph p1, Paragraph p2)
    {
        // Compare style name and alignment. Add more properties here if stricter comparison is needed.
        return string.Equals(p1.ParagraphFormat.StyleName, p2.ParagraphFormat.StyleName, StringComparison.Ordinal) &&
               p1.ParagraphFormat.Alignment == p2.ParagraphFormat.Alignment;
    }
}
