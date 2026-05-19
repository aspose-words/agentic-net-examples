using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;          // Needed for Footnote and FootnoteType
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Build a sample document with narrow page width to force wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        doc.FirstSection.PageSetup.PageWidth = 300;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Write a paragraph with long words that can be hyphenated.
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Insert a footnote containing the same long text.
        builder.InsertFootnote(FootnoteType.Footnote,
            "extraordinarycharacteristically internationalization communication");

        // Save the original layout (hyphenation applied to footnote as well).
        const string beforePdf = "FootnoteHyphenation_Before.pdf";
        doc.Save(beforePdf, SaveFormat.Pdf);
        if (!File.Exists(beforePdf))
            throw new InvalidOperationException("Failed to create the before PDF.");

        // Disable hyphenation for footnote paragraphs only.
        foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        {
            Paragraph? firstParagraph = footnote.FirstParagraph;
            if (firstParagraph != null)
                firstParagraph.ParagraphFormat.SuppressAutoHyphens = true;
        }

        // Save the modified layout.
        const string afterPdf = "FootnoteHyphenation_After.pdf";
        doc.Save(afterPdf, SaveFormat.Pdf);
        if (!File.Exists(afterPdf))
            throw new InvalidOperationException("Failed to create the after PDF.");

        // Simple comparison: file size difference indicates layout change.
        long beforeSize = new FileInfo(beforePdf).Length;
        long afterSize = new FileInfo(afterPdf).Length;

        Console.WriteLine($"Before PDF size: {beforeSize} bytes");
        Console.WriteLine($"After PDF size:  {afterSize} bytes");
        Console.WriteLine(afterSize != beforeSize
            ? "Footnote hyphenation was successfully disabled."
            : "No change detected in footnote hyphenation.");
    }
}
