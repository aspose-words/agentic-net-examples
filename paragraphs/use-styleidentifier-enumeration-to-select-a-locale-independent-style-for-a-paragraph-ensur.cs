using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder for easy paragraph insertion and formatting.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a built‑in heading style using the locale‑independent StyleIdentifier.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("This paragraph uses the Heading 1 style.");

        // Apply the normal style (also locale‑independent) to the next paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This paragraph uses the Normal style.");

        // Save the document to the local file system.
        string outputPath = "StyledParagraph.docx";
        doc.Save(outputPath);
    }
}
