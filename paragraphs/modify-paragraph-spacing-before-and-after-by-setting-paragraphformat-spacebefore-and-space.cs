using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with sample text.
        builder.Writeln("This is a sample paragraph.");

        // Retrieve the paragraph we just added.
        Paragraph paragraph = doc.LastSection.Body.LastParagraph;

        // Set spacing before and after (measured in points).
        paragraph.ParagraphFormat.SpaceBefore = 12; // 12 points before the paragraph
        paragraph.ParagraphFormat.SpaceAfter = 12;  // 12 points after the paragraph

        // Save the modified document.
        doc.Save("ModifiedParagraphSpacing.docx");
    }
}
