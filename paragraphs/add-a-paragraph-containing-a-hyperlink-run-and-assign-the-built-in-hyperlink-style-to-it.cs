using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new paragraph.
        builder.Writeln();

        // Apply the built‑in Hyperlink character style to the upcoming text.
        builder.Font.StyleIdentifier = StyleIdentifier.Hyperlink;

        // Insert a hyperlink run with display text and target URL.
        builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);

        // End the paragraph.
        builder.Writeln();

        // Save the document to a file.
        doc.Save("HyperlinkParagraph.docx");
    }
}
