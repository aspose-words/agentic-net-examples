using System;
using Aspose.Words;
using static Aspose.Words.Hyphenation; // Hyphenation is a static class, not a namespace

class HyphenateRtf
{
    static void Main()
    {
        // Load the source RTF document.
        Document doc = new Document("Input.rtf");

        // Optional: register a hyphenation dictionary for the document language (e.g., English US).
        // Hyphenation.RegisterDictionary("en-US", "hyph_en_US.dic");
        // Using the static import above, we can call RegisterDictionary directly:
        // RegisterDictionary("en-US", "hyph_en_US.dic");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Configure additional hyphenation settings (values are illustrative).
        doc.HyphenationOptions.HyphenationZone = 720;          // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;    // Max two consecutive hyphenated lines.
        doc.HyphenationOptions.HyphenateCaps = true;         // Hyphenate words in all caps.

        // Save the document back to RTF. Aspose.Words inserts optional hyphen characters
        // (ControlChar.OptionalHyphenChar) at the appropriate break points.
        doc.Save("Output.rtf");
    }
}
