using System;
using Aspose.Words;

public class MergeParagraphsExample
{
    public static void Main()
    {
        // Create a new document and a builder for inserting content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph 1 – style "Normal", left aligned.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Writeln("First paragraph with normal style.");

        // Paragraph 2 – same formatting as paragraph 1 (should be merged).
        builder.Writeln("Second paragraph with the same formatting.");

        // Paragraph 3 – different alignment (won't be merged with previous).
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Writeln("Third paragraph centered.");

        // Paragraph 4 – same formatting as paragraph 3 (should be merged).
        builder.Writeln("Fourth paragraph also centered.");

        // Paragraph 5 – different style (won't be merged).
        builder.ParagraphFormat.StyleName = "Heading 1";
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.Writeln("Heading paragraph.");

        // Save the original document for reference.
        doc.Save("Original.docx");

        // Merge consecutive paragraphs that share identical formatting.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        int i = 0;
        while (i < paragraphs.Count - 1)
        {
            Paragraph current = paragraphs[i];
            Paragraph next = paragraphs[i + 1];

            // Compare formatting: style name and alignment.
            bool sameStyle = string.Equals(current.ParagraphFormat.StyleName, next.ParagraphFormat.StyleName, StringComparison.Ordinal);
            bool sameAlignment = current.ParagraphFormat.Alignment == next.ParagraphFormat.Alignment;

            if (sameStyle && sameAlignment)
            {
                // Move all child nodes (runs, fields, etc.) from the next paragraph to the current one.
                while (next.HasChildNodes)
                {
                    Node child = next.FirstChild;
                    next.RemoveChild(child);
                    current.AppendChild(child);
                }

                // Remove the now empty next paragraph.
                next.Remove();

                // Do not increment i to check the new next paragraph against the current one.
            }
            else
            {
                i++; // Move to the next pair.
            }
        }

        // Save the merged document.
        doc.Save("Merged.docx");
    }
}
