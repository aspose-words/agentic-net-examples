using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text.
        builder.Writeln("This is a Heading 2 paragraph.");

        // Apply the built‑in "Heading 2" style to the current paragraph.
        builder.CurrentParagraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

        // Adjust spacing before and after the paragraph (points).
        builder.CurrentParagraph.ParagraphFormat.SpaceBefore = 12; // 12 points before
        builder.CurrentParagraph.ParagraphFormat.SpaceAfter = 6;   // 6 points after

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
