using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Path to the folder that contains the hyphenation dictionary file.
        // The dictionary file must be in OpenOffice format (e.g., hyph_en_US.dic).
        string dataDir = @"C:\HyphenationData\"; // <-- adjust to your environment
        string dictionaryPath = Path.Combine(dataDir, "hyph_en_US.dic");

        // Register the English (US) hyphenation dictionary.
        // This enables Aspose.Words to hyphenate words according to the same rules as Microsoft Word.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Verify that the dictionary has been registered successfully.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the en‑US hyphenation dictionary.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: limit the number of consecutive hyphenated lines (0 = no limit).
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 0;

        // Optional: set the hyphenation zone (distance from the right margin where hyphenation is suppressed).
        // 360 = 0.25 inch (default). Increase to reduce hyphenation near the margin.
        doc.HyphenationOptions.HyphenationZone = 360;

        // Write a paragraph that contains a long word which will be hyphenated by Word.
        // The word "characteristically" is long enough to be split when the line width is limited.
        builder.Font.Size = 24;
        builder.Writeln(
            "The word characteristically demonstrates how automatic hyphenation works when the line width is constrained.");

        // Force a narrow page width to make hyphenation visible.
        // This mimics the layout conditions where Word would insert hyphens.
        doc.FirstSection.PageSetup.PageWidth = 300;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Update the layout so that hyphenation is applied.
        doc.UpdatePageLayout();

        // Save the document. The resulting file will show hyphenated lines just like Microsoft Word.
        string outputPath = Path.Combine(dataDir, "HyphenatedDocument.docx");
        doc.Save(outputPath);
    }
}
