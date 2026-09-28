using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will have custom spacing before and after.");

        // Retrieve the paragraph that was just added.
        Paragraph paragraph = builder.CurrentParagraph;

        // Set spacing before and after (values are in points).
        paragraph.ParagraphFormat.SpaceBefore = 12f; // 12 points before the paragraph
        paragraph.ParagraphFormat.SpaceAfter = 6f;   // 6 points after the paragraph

        // Save the document to a file.
        doc.Save("ParagraphSpacing.docx");
    }
}
