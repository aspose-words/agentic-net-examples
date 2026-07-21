using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationReport
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        if (!File.Exists(dictFileName))
        {
            // The dictionary format: first line is "UTF-8", subsequent lines are word=hyphenation-points.
            // Include a few long words to demonstrate hyphenation.
            File.WriteAllText(dictFileName,
                "UTF-8\n" +
                "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
                "internationalization=in-ter-na-tion-al-i-za-tion\n" +
                "communication=com-mu-ni-ca-tion\n");
        }

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: tweak hyphenation settings.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.17 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // ----- Add a heading (hyphenation disabled) -----
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        // Suppress hyphenation for this heading paragraph.
        builder.ParagraphFormat.SuppressAutoHyphens = true;
        builder.Writeln("Extraordinarycharacteristically Internationalization Communication");

        // Reset paragraph formatting for body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.SuppressAutoHyphens = false;

        // ----- Add body paragraphs (hyphenation enabled) -----
        builder.Font.Size = 12;
        builder.Writeln(
            "This report demonstrates how automatic hyphenation works in a document. " +
            "When the lines are too long to fit within the page margins, words may be split " +
            "according to the hyphenation dictionary. The following paragraph contains several " +
            "long words such as extraordinarycharacteristically, internationalization, and " +
            "communication that will be hyphenated if needed.");

        // Add another body paragraph to ensure multiple lines.
        builder.Writeln(
            "Hyphenation improves the visual appearance of justified text and reduces ragged " +
            "edges. By disabling hyphenation for headings, we keep headings clean and " +
            "readable, while still benefiting from hyphenation in the body of the report.");

        // Save the document to PDF to render hyphenation.
        const string outputFile = "HyphenationReport.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException($"Expected output file '{outputFile}' was not created.");
    }
}
