using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in "Heading 2" style to the paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

        // Adjust spacing before and after the paragraph (values are in points).
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before
        builder.ParagraphFormat.SpaceAfter = 6;   // 6 points after

        // Add some text that will use the Heading 2 style.
        builder.Writeln("Sample Heading 2");

        // Ensure the output directory exists.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Heading2_Styled.docx");
        doc.Save(outputPath);
    }
}
