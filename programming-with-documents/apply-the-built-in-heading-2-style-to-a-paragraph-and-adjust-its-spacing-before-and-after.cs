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

        // Set the paragraph style to the built‑in "Heading 2".
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

        // Adjust spacing before and after the paragraph (in points).
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before
        builder.ParagraphFormat.SpaceAfter = 6;   // 6 points after

        // Write the paragraph text.
        builder.Writeln("This paragraph uses the built‑in Heading 2 style.");

        // Save the document to a file.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
