using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Heading 2 style to the current paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

        // Set custom spacing before and after the paragraph (in points).
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before
        builder.ParagraphFormat.SpaceAfter = 12;  // 12 points after

        // Add some sample text to the styled paragraph.
        builder.Writeln("This is a Heading 2 paragraph with custom spacing.");

        // Save the document to the local file system.
        doc.Save("Heading2_Styled.docx");
    }
}
