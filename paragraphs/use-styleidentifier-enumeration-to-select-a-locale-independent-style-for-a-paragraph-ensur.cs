using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add paragraphs.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a locale‑independent style using StyleIdentifier.
        // This ensures the same formatting regardless of the document language.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("This paragraph uses the Heading1 style.");

        // Change to another built‑in style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This paragraph uses the Normal style.");

        // Apply a third style for demonstration.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Quote;
        builder.Writeln("This paragraph uses the Quote style.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledParagraphs.docx");
        doc.Save(outputPath);
    }
}
