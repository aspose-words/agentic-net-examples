using System;
using Aspose.Words;

namespace ParagraphSpacingExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set custom spacing before and after the paragraph (in points).
            builder.ParagraphFormat.SpaceBefore = 12; // 12 points before the paragraph
            builder.ParagraphFormat.SpaceAfter = 24;  // 24 points after the paragraph

            // Insert a paragraph that will use the spacing settings above.
            builder.Writeln("This paragraph has custom spacing before and after.");

            // Save the document to the current directory.
            doc.Save("ParagraphSpacingOutput.docx");
        }
    }
}
