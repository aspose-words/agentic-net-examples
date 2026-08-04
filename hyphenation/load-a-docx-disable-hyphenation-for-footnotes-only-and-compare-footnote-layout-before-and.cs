using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (en-US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n");

        // Register the dictionary for the "en-US" language.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document locale so the hyphenation dictionary is applied.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Use the default hyphenation zone (no need to set it to 0, which is invalid).
        // doc.HyphenationOptions.HyphenationZone = 360; // optional

        // Write a paragraph containing a long word that can be hyphenated.
        const string longWord = "extraordinarycharacteristically";
        builder.Font.Size = 12;
        builder.Writeln($"This paragraph contains a long word: {longWord}.");

        // Insert a footnote with the same long word.
        builder.InsertFootnote(FootnoteType.Footnote, $"Footnote with the same long word: {longWord}.");

        // Save the document before disabling hyphenation in footnotes.
        const string beforePdf = "FootnoteHyphenation_Before.pdf";
        doc.Save(beforePdf, SaveFormat.Pdf);
        if (!File.Exists(beforePdf))
            throw new InvalidOperationException("Failed to create the before‑hyphenation PDF.");

        // Disable hyphenation for footnote paragraphs only.
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote footnote in footnotes)
        {
            Paragraph? para = footnote.FirstParagraph;
            if (para != null)
                para.ParagraphFormat.SuppressAutoHyphens = true;
        }

        // Save the document after disabling hyphenation in footnotes.
        const string afterPdf = "FootnoteHyphenation_After.pdf";
        doc.Save(afterPdf, SaveFormat.Pdf);
        if (!File.Exists(afterPdf))
            throw new InvalidOperationException("Failed to create the after‑hyphenation PDF.");

        // Output the names of the generated files.
        Console.WriteLine($"Created PDFs: {beforePdf}, {afterPdf}");
    }
}
