using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder for convenient content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph formatting: justified alignment and enable word wrap.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        builder.ParagraphFormat.WordWrap = true; // Ensures whole words are wrapped, not mid‑word.

        // Add sample text that will demonstrate the justification and wrapping.
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                        "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Save the document to the local file system.
        string outputPath = "JustifiedParagraph.docx";
        doc.Save(outputPath);
    }
}
