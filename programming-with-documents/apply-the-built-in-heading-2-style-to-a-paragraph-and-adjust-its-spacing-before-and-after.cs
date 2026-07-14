using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Heading 2 style to the current paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

        // Set spacing before and after the paragraph (points).
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before
        builder.ParagraphFormat.SpaceAfter = 12;  // 12 points after

        // Add some text that will be formatted with Heading 2.
        builder.Writeln("Sample Heading 2");

        // Save the document.
        doc.Save("Heading2Spacing.docx");
    }
}
