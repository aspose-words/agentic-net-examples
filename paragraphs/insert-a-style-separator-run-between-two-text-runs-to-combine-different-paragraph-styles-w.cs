using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First part of the paragraph with Heading1 paragraph style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("This text uses Heading 1 style. ");

        // Insert a style separator run.
        builder.InsertStyleSeparator();

        // Second part of the same paragraph with a character style.
        // Use DefaultParagraphFont (a built‑in character style) to revert to normal formatting.
        builder.Font.StyleIdentifier = StyleIdentifier.DefaultParagraphFont;
        builder.Write("This text uses Normal style.");

        // Save the document.
        doc.Save("StyleSeparatorExample.docx");
    }
}
