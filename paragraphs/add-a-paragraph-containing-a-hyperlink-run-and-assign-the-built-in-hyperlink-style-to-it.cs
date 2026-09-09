using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fonts;

public class HyperlinkParagraphExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new paragraph.
        builder.Writeln("Paragraph containing a hyperlink:");

        // Apply the built‑in Hyperlink character style to the upcoming text.
        builder.Font.StyleIdentifier = StyleIdentifier.Hyperlink;

        // Insert the hyperlink field. The display text is "Visit Aspose".
        builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);

        // Clear the formatting so subsequent text is not affected.
        builder.Font.ClearFormatting();

        // End the paragraph.
        builder.Writeln();

        // Save the document to the same folder as the executable.
        doc.Save("HyperlinkParagraph.docx");
    }
}
