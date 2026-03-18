using System;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables;

class HyphenationAndLineSpacingExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample text that will benefit from hyphenation.
        builder.Font.Size = 24;
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                        "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                        "Praesent commodo cursus magna, vel scelerisque nisl consectetur et.");

        // Enable automatic hyphenation and configure its options.
        doc.HyphenationOptions.AutoHyphenation = true;          // Turn on hyphenation.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;     // Limit consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch zone from the right margin.
        doc.HyphenationOptions.HyphenateCaps = true;           // Hyphenate all‑caps words.

        // Adjust line spacing for all paragraphs to improve readability.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Use multiple line spacing (e.g., 1.5 lines).
            para.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
            para.ParagraphFormat.LineSpacing = 1.5; // 150% of the normal line height.
        }

        // Save the resulting document.
        doc.Save("HyphenatedAndSpaced.docx");
    }
}
