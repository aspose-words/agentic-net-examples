using System;
using Aspose.Words;
using Aspose.Words.Tables;

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

            // Move the builder's cursor to the third paragraph (index 2).
            // The second parameter (characterIndex) is set to 0 to position at the start.
            builder.MoveToParagraph(2, 0);

            // Apply formatting to the paragraph at the current cursor position.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.ParagraphFormat.StyleName = "Heading 1";

            // Optionally, add additional text to the formatted paragraph.
            builder.Writeln(" - This paragraph has been centered and styled as Heading 1.");

            // Save the document to the output file.
            doc.Save("ParagraphNavigationOutput.docx");
        }
    }
}
