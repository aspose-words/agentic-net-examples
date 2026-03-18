using System;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with long justified text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        builder.Writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut " +
            "labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco " +
            "laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit " +
            "in voluptate velit esse cillum dolore eu fugiat nulla pariatur.");

        // Enable automatic hyphenation so words can be split across lines.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Reduce the hyphenation zone to avoid hyphenating words too close to the right margin.
        doc.HyphenationOptions.HyphenationZone = 360; // 0.25 inch (default)

        // Change the document-wide justification mode to compress spacing.
        // This reduces the large gaps that can appear when a justified paragraph contains hyphenated lines.
        doc.JustificationMode = JustificationMode.Compress;

        // Optional: prevent justification of lines that end with a soft line break (Shift+Enter).
        // This further reduces unwanted spacing in certain cases.
        doc.CompatibilityOptions.DoNotExpandShiftReturn = true;

        // Save the result.
        doc.Save("AdjustedJustification.docx");
    }
}
