using System;
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
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Write a paragraph with a long word that can be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln(
            "This paragraph contains a long word that will be hyphenated: extraordinarycharacteristically.");

        // Insert a footnote that contains the same long word.
        builder.InsertFootnote(FootnoteType.Footnote,
            "Footnote text with the same long word: extraordinarycharacteristically.");

        // Save the document before disabling hyphenation in footnotes.
        const string beforePath = "FootnoteHyphenation_Before.pdf";
        doc.Save(beforePath, SaveFormat.Pdf);
        if (!File.Exists(beforePath))
            throw new InvalidOperationException("Failed to create the 'before' PDF.");

        // Disable hyphenation for all footnote paragraphs only.
        foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        {
            // Each footnote contains a paragraph as its first child.
            if (footnote.FirstParagraph?.ParagraphFormat != null)
                footnote.FirstParagraph.ParagraphFormat.SuppressAutoHyphens = true;
        }

        // Verify that the suppression flag is set.
        foreach (Footnote footnote in doc.GetChildNodes(NodeType.Footnote, true))
        {
            if (footnote.FirstParagraph?.ParagraphFormat?.SuppressAutoHyphens != true)
                throw new InvalidOperationException("Failed to suppress hyphenation on a footnote paragraph.");
        }

        // Save the document after disabling hyphenation in footnotes.
        const string afterPath = "FootnoteHyphenation_After.pdf";
        doc.Save(afterPath, SaveFormat.Pdf);
        if (!File.Exists(afterPath))
            throw new InvalidOperationException("Failed to create the 'after' PDF.");

        // Simple comparison: report file sizes (they will differ if hyphenation changed layout).
        long beforeSize = new FileInfo(beforePath).Length;
        long afterSize = new FileInfo(afterPath).Length;

        Console.WriteLine($"Before disabling footnote hyphenation: {beforeSize} bytes");
        Console.WriteLine($"After disabling footnote hyphenation : {afterSize} bytes");
        Console.WriteLine("Hyphenation suppression applied to footnotes only.");
    }
}
