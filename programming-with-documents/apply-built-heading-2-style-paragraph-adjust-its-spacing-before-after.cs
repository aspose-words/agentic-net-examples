using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Heading 2 style to the current paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

        // Adjust spacing before and after the paragraph (values are in points).
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before the paragraph
        builder.ParagraphFormat.SpaceAfter = 6;   // 6 points after the paragraph

        // Add some text to the styled paragraph.
        builder.Writeln("This is a Heading 2 paragraph with custom spacing.");

        // Save the document to a file.
        doc.Save("Heading2_Styled.docx");
    }
}
