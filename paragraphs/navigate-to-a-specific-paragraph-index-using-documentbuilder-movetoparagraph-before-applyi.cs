using System;
using Aspose.Words;

namespace ParagraphNavigationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add several paragraphs.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Paragraph 0: This is the first paragraph.");
            builder.Writeln("Paragraph 1: This is the second paragraph.");
            builder.Writeln("Paragraph 2: This is the third paragraph.");
            builder.Writeln("Paragraph 3: This is the fourth paragraph.");

            // Move the builder's cursor to the third paragraph (index 2) at the start of the paragraph.
            // The second parameter (characterIndex) is set to 0 to position at the beginning.
            builder.MoveToParagraph(2, 0);

            // Apply formatting changes to the selected paragraph.
            // For example, center-align the paragraph text.
            builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Optionally, add additional text to the formatted paragraph.
            builder.Write(" (center-aligned)");

            // Save the modified document to a file.
            doc.Save("Output.docx");
        }
    }
}
