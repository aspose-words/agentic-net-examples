using System;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Turn on automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // ---------- First section (hyphenation disabled) ----------
        builder.Writeln("First section paragraph. This paragraph will not be hyphenated even though the document has hyphenation enabled.");
        // Suppress hyphenation for this paragraph.
        builder.CurrentParagraph.ParagraphFormat.SuppressAutoHyphens = true;

        // Insert a new section that will contain the paragraph with hyphenation.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Target paragraph (hyphenation enabled) ----------
        builder.Writeln(
            "This is a long paragraph that will demonstrate automatic hyphenation. " +
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        // Ensure hyphenation is NOT suppressed for this paragraph (default is false, set explicitly for clarity).
        builder.CurrentParagraph.ParagraphFormat.SuppressAutoHyphens = false;

        // Insert another section after the target paragraph.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Last section (hyphenation disabled) ----------
        builder.Writeln("Last section paragraph. Hyphenation is disabled here as well.");
        // Suppress hyphenation for this paragraph.
        builder.CurrentParagraph.ParagraphFormat.SuppressAutoHyphens = true;

        // Save the document.
        doc.Save("HyphenationExample.docx");
    }
}
