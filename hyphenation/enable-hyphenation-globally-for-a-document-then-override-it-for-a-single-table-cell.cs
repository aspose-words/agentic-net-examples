using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal English hyphenation dictionary locally.
        const string dictionaryPath = "hyph_en_US.dic";
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create a new document and configure a narrow page to force line wrapping.
        Document doc = new Document();
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720;
        doc.HyphenationOptions.HyphenateCaps = true;

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;

        // Paragraph that will be hyphenated automatically.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Insert a table with a single cell where hyphenation is suppressed.
        builder.StartTable();
        builder.InsertCell();

        // Suppress hyphenation for the paragraph inside this cell.
        Paragraph cellParagraph = builder.CurrentParagraph;
        cellParagraph.ParagraphFormat.SuppressAutoHyphens = true;

        // Text that would normally be hyphenated, but will not be in this cell.
        builder.Writeln("extraordinarycharacteristically");

        builder.EndRow();
        builder.EndTable();

        // Save the document as PDF to visualize hyphenation.
        const string outputPath = "HyphenationExample.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF output was not created.");
    }
}
