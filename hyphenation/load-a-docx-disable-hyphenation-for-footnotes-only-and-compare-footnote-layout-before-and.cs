using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (en-US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a sample document with narrow page width to force line wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page setup.
        doc.FirstSection.PageSetup.PageWidth = 200;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the document locale to English (United States).
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Add a paragraph containing a long word that can be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln(
            "This paragraph contains a long word extraordinarycharacteristically that should be hyphenated when the line is too short.");

        // Insert a footnote that also contains the same long word.
        builder.InsertFootnote(FootnoteType.Footnote,
            "Footnote with extraordinarycharacteristically word that may be hyphenated.");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document before disabling hyphenation in footnotes.
        const string beforePath = "FootnoteHyphenation_Before.pdf";
        doc.Save(beforePath, SaveFormat.Pdf);
        if (!File.Exists(beforePath))
            throw new InvalidOperationException("Failed to create the 'before' PDF.");

        // Disable hyphenation only for footnote paragraphs.
        foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        {
            if (footnote.FirstParagraph != null)
                footnote.FirstParagraph.ParagraphFormat.SuppressAutoHyphens = true;
        }

        // Save the document after disabling footnote hyphenation.
        const string afterPath = "FootnoteHyphenation_After.pdf";
        doc.Save(afterPath, SaveFormat.Pdf);
        if (!File.Exists(afterPath))
            throw new InvalidOperationException("Failed to create the 'after' PDF.");

        // Compare the resulting file sizes as a simple indication of layout change.
        long beforeSize = new FileInfo(beforePath).Length;
        long afterSize = new FileInfo(afterPath).Length;

        Console.WriteLine($"Before PDF size: {beforeSize} bytes");
        Console.WriteLine($"After PDF size : {afterSize} bytes");
    }
}
