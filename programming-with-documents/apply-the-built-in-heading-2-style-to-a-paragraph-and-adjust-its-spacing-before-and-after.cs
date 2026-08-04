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

        // Set custom spacing before and after the paragraph (in points).
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before
        builder.ParagraphFormat.SpaceAfter = 12;  // 12 points after

        // Write the paragraph text.
        builder.Writeln("This paragraph uses Heading 2 style with custom spacing.");

        // Save the document to a file in the current directory.
        doc.Save("Heading2_Styled.docx");
    }
}
