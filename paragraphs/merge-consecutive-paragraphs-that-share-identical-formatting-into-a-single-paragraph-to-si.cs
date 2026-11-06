using System;
using Aspose.Words;

public class Program
{
    // Determines whether two paragraphs have identical formatting.
    // For simplicity we compare the paragraph style and the font of the first run in each paragraph.
    private static bool HaveSameFormatting(Paragraph p1, Paragraph p2)
    {
        // Compare paragraph styles.
        if (p1.ParagraphFormat.StyleIdentifier != p2.ParagraphFormat.StyleIdentifier ||
            p1.ParagraphFormat.StyleName != p2.ParagraphFormat.StyleName)
            return false;

        // Get the first run of each paragraph (if any).
        Run run1 = p1.GetChildNodes(NodeType.Run, true).Count > 0
            ? (Run)p1.GetChildNodes(NodeType.Run, true)[0]
            : null;
        Run run2 = p2.GetChildNodes(NodeType.Run, true).Count > 0
            ? (Run)p2.GetChildNodes(NodeType.Run, true)[0]
            : null;

        // If both paragraphs have no runs, consider them equal in font formatting.
        if (run1 == null && run2 == null)
            return true;
        if (run1 == null || run2 == null)
            return false;

        // Compare font properties.
        return run1.Font.Name == run2.Font.Name &&
               run1.Font.Size == run2.Font.Size &&
               run1.Font.Bold == run2.Font.Bold &&
               run1.Font.Italic == run2.Font.Italic;
    }

    public static void Main()
    {
        // Create a new document and add sample paragraphs with varying formatting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph 1 – Bold Arial 12pt
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;
        builder.Font.Bold = true;
        builder.Writeln("First paragraph, bold Arial 12pt.");

        // Paragraph 2 – Same formatting as Paragraph 1
        builder.Writeln("Second paragraph, same formatting as first.");

        // Paragraph 3 – Regular Times New Roman 14pt
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 14;
        builder.Font.Bold = false;
        builder.Writeln("Third paragraph, regular Times New Roman 14pt.");

        // Paragraph 4 – Same formatting as Paragraph 3
        builder.Writeln("Fourth paragraph, same formatting as third.");

        // Merge consecutive paragraphs that share identical formatting.
        Body body = doc.FirstSection.Body;

        // Iterate backwards to safely remove paragraphs while traversing.
        for (int i = body.Paragraphs.Count - 2; i >= 0; i--)
        {
            Paragraph current = body.Paragraphs[i];
            Paragraph next = body.Paragraphs[i + 1];

            if (HaveSameFormatting(current, next))
            {
                // Retrieve text of the next paragraph without the trailing paragraph mark.
                string nextText = next.GetText();
                if (nextText.EndsWith("\r"))
                    nextText = nextText.Substring(0, nextText.Length - 1);

                // Append a space and the next paragraph's text to the current paragraph.
                Run run = new Run(doc, " " + nextText);
                current.AppendChild(run);

                // Remove the now-merged next paragraph.
                next.Remove();
            }
        }

        // Save the resulting document.
        doc.Save("MergedParagraphs.docx");
    }
}
