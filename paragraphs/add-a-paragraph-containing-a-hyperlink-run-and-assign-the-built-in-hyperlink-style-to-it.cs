using System;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Hyperlink character style to the text that will be inserted.
        // Hyperlink is a character style, so we set it on the Font, not on the Paragraph.
        builder.Font.StyleIdentifier = StyleIdentifier.Hyperlink;

        // Insert a hyperlink field into the current paragraph.
        builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);

        // Reset font formatting for any subsequent text.
        builder.Font.ClearFormatting();

        // Save the document.
        doc.Save("HyperlinkParagraph.docx");
    }
}
