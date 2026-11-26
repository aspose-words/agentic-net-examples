using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph whose spacing will be modified.");

        // Retrieve the first paragraph in the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Set spacing before and after the paragraph (values are in points).
        paragraph.ParagraphFormat.SpaceBefore = 12; // 12 points before the paragraph.
        paragraph.ParagraphFormat.SpaceAfter = 24;  // 24 points after the paragraph.

        // Save the document to disk.
        doc.Save("ModifiedParagraphSpacing.docx");
    }
}
