using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a common formatting for the first group of paragraphs.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Font.Size = 12;
        builder.Font.Color = Color.Black;

        // These two paragraphs have identical formatting and should be merged.
        builder.Writeln("Paragraph 1 – same format.");
        builder.Writeln("Paragraph 2 – same format.");

        // Change formatting – this paragraph must stay separate.
        builder.Font.Color = Color.Red;
        builder.Writeln("Paragraph 3 – different format.");

        // Revert to the original formatting – not consecutive with the first group,
        // so it will remain a separate paragraph.
        builder.Font.Color = Color.Black;
        builder.Writeln("Paragraph 4 – same as first group.");

        // Merge consecutive paragraphs that share identical formatting.
        MergeConsecutiveParagraphs(doc);

        // Save the resulting document.
        doc.Save("MergedParagraphs.docx");
    }

    // Returns true if two paragraphs have the same formatting that matters for merging.
    private static bool HaveSameFormatting(Paragraph p1, Paragraph p2)
    {
        // Compare style identifier and name. If both are equal, we treat the formatting as identical.
        // Additional properties (alignment, indents, etc.) can be added here if needed.
        return p1.ParagraphFormat.StyleIdentifier == p2.ParagraphFormat.StyleIdentifier &&
               string.Equals(p1.ParagraphFormat.StyleName, p2.ParagraphFormat.StyleName, StringComparison.Ordinal);
    }

    // Merges each paragraph with the previous one when their formatting matches.
    private static void MergeConsecutiveParagraphs(Document doc)
    {
        // Work on the body of the first section.
        Body body = doc.FirstSection.Body;
        // ParagraphCollection provides indexed access.
        ParagraphCollection paragraphs = body.Paragraphs;

        int i = 1; // Start from the second paragraph.
        while (i < paragraphs.Count)
        {
            Paragraph previous = paragraphs[i - 1];
            Paragraph current = paragraphs[i];

            if (HaveSameFormatting(previous, current))
            {
                // Ensure the previous paragraph has at least one run to receive the text.
                if (previous.Runs.Count == 0)
                {
                    Run emptyRun = new Run(doc);
                    previous.AppendChild(emptyRun);
                }

                // Append the text of the current paragraph (without its terminating paragraph break)
                // to the last run of the previous paragraph.
                Run lastRun = (Run)previous.Runs[previous.Runs.Count - 1];
                string currentText = current.GetText();
                // Remove the final paragraph break character (\r) if present.
                if (currentText.EndsWith("\r"))
                    currentText = currentText.Substring(0, currentText.Length - 1);
                lastRun.Text += currentText;

                // Remove the current paragraph from the document.
                current.Remove();
                // Do not increment i because the next paragraph shifts into the current index.
            }
            else
            {
                i++; // Formatting differs – move to the next pair.
            }
        }
    }
}
