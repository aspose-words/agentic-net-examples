using System;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Hyperlink character style to the upcoming text.
        builder.Font.StyleIdentifier = StyleIdentifier.Hyperlink;

        // Insert a hyperlink field. The display text will be styled with the Hyperlink style.
        builder.InsertHyperlink("Aspose.Words", "https://www.aspose.com/words", false);

        // End the paragraph.
        builder.Writeln();

        // Save the document to the local file system.
        doc.Save("HyperlinkParagraph.docx");
    }
}
