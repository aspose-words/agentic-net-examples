using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which simplifies building the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a locale‑independent style to the current paragraph using StyleIdentifier.
        // This ensures the same formatting regardless of the document language.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("This paragraph uses the Heading 1 style.");

        // Change to another built‑in style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This paragraph uses the Normal style.");

        // Save the document.
        string outputPath = "StyledParagraph.docx";
        doc.Save(outputPath);
    }
}
