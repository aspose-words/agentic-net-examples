using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // HTML fragment that contains a long paragraph.
        string html = "<p>This is a demonstration of automatic hyphenation in a long paragraph that will be broken across lines to show hyphen insertion.</p>";

        // Insert the HTML into the document.
        builder.InsertHtml(html);

        // Turn on automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Max two consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenateCaps = true;         // Hyphenate words in all caps.

        // Save the document as HTML. Hyphenated words will contain soft‑hyphen characters (U+00AD) which render as “&shy;” in HTML.
        doc.Save("HyphenatedDocument.html", SaveFormat.Html);
    }
}
