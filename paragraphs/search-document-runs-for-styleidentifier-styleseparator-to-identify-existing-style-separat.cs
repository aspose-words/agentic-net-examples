using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace StyleSeparatorSearch
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a paragraph that contains a style separator.
            // First part uses Heading1 style.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Write("This is heading text. ");

            // Insert a style separator so the next text can have a different style
            // but remain on the same line.
            builder.InsertStyleSeparator();

            // Second part uses Quote style.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
            builder.Write("This is a quoted text. ");

            // End the paragraph.
            builder.InsertParagraph();

            // Add a normal paragraph without a style separator for comparison.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("A regular paragraph without a style separator.");

            // Save the document (optional, just to visualize the result).
            doc.Save("StyleSeparatorSearch.docx");

            // Search for paragraphs whose break is a style separator.
            int separatorCount = 0;
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.BreakIsStyleSeparator)
                {
                    separatorCount++;
                    // Output the text of the paragraph that contains the style separator.
                    Console.WriteLine($"Paragraph with style separator found: \"{para.GetText().Trim()}\"");
                }
            }

            Console.WriteLine($"Total paragraphs with style separators: {separatorCount}");
        }
    }
}
