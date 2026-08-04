using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fonts;
using System.Drawing;

public class HyperlinkParagraphExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new paragraph.
        builder.InsertParagraph();

        // Apply the built‑in Hyperlink character style to the upcoming text.
        // This style sets the typical blue color and underline for hyperlinks.
        builder.Font.StyleIdentifier = StyleIdentifier.Hyperlink;

        // Insert a hyperlink field with display text and URL.
        // The third argument (false) indicates that the second parameter is a URL, not a bookmark.
        builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);

        // End the paragraph.
        builder.Writeln();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HyperlinkParagraph.docx");
        doc.Save(outputPath);
    }
}
