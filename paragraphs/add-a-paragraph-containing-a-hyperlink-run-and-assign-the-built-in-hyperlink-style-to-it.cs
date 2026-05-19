using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply the built‑in Hyperlink character style to the upcoming run.
        builder.Font.StyleIdentifier = StyleIdentifier.Hyperlink;

        // Insert a hyperlink field. The Hyperlink style will give it the default
        // blue color and underline, but we can also set them explicitly if desired.
        builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);

        // End the paragraph.
        builder.Writeln();

        // Save the document.
        doc.Save("HyperlinkParagraph.docx");
    }
}
