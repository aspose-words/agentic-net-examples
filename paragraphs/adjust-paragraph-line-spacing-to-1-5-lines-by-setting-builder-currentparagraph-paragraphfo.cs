using System;
using Aspose.Words;

namespace ParagraphLineSpacingExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a new paragraph and obtain a reference to it.
            // The builder's CurrentParagraph now points to this paragraph.
            builder.InsertParagraph();

            // Set the line spacing of the current paragraph to 1.5 lines.
            // Note: When LineSpacingRule is not set, the value is interpreted as points.
            // For demonstration purposes we directly assign 1.5 as requested.
            builder.CurrentParagraph.ParagraphFormat.LineSpacing = 1.5;

            // Add some text to the paragraph.
            builder.Writeln("This paragraph has a line spacing of 1.5 lines.");

            // Save the document to the local file system.
            doc.Save("ParagraphLineSpacing.docx");
        }
    }
}
