using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class MergeParagraphsExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs. Some of them share the same formatting (Normal style).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("First paragraph with Normal style.");
        builder.Writeln("Second paragraph with Normal style."); // Same formatting as previous.

        // Change formatting to a different style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("A heading paragraph."); // Different formatting.

        // Return to Normal style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Third paragraph with Normal style."); // Same formatting as the first two, but not consecutive.

        // Merge consecutive paragraphs that have identical formatting.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Iterate through the collection, comparing each paragraph with its predecessor.
        for (int i = 1; i < paragraphs.Count; i++)
        {
            Paragraph previous = paragraphs[i - 1];
            Paragraph current = paragraphs[i];

            // Determine if the two paragraphs share the same formatting.
            bool sameStyle = previous.ParagraphFormat.StyleIdentifier == current.ParagraphFormat.StyleIdentifier;
            bool sameAlignment = previous.ParagraphFormat.Alignment == current.ParagraphFormat.Alignment;

            // Add more property comparisons here if stricter matching is required.

            if (sameStyle && sameAlignment)
            {
                // Append the text of the current paragraph to the previous one.
                // Trim the paragraph break characters that GetText() appends.
                string currentText = current.GetText().TrimEnd('\r', '\n', '\x0c');
                Run mergedRun = new Run(doc, currentText);
                previous.AppendChild(mergedRun);

                // Remove the now-merged current paragraph.
                current.Remove();

                // After removal, the collection shrinks, so stay at the same index.
                i--;
            }
        }

        // Save the resulting document.
        doc.Save("MergedParagraphs.docx");
    }
}
