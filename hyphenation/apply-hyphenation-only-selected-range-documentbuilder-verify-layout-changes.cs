using System;
using Aspose.Words;
using Aspose.Words.Layout;

class HyphenationRangeExample
{
    static void Main()
    {
        // Create a new blank document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a long line of text that will need hyphenation when it reaches the right margin.
        string longText = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.";

        // First paragraph – we will suppress hyphenation for this range.
        builder.ParagraphFormat.SuppressAutoHyphens = true; // Exempt this paragraph from hyphenation.
        builder.Writeln(longText);

        // Second paragraph – hyphenation will be allowed.
        builder.ParagraphFormat.SuppressAutoHyphens = false; // Ensure hyphenation is enabled.
        builder.Writeln(longText);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: configure additional hyphenation settings.
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Rebuild the layout so that hyphenation takes effect.
        doc.UpdatePageLayout();

        // Save the document – the first paragraph will have no hyphens, the second will.
        doc.Save("HyphenationRange.docx");
    }
}
