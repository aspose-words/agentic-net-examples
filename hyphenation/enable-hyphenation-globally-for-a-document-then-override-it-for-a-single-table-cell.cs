using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables; // Needed for the Table class

public class HyphenationExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a minimal hyphenation dictionary for English (US).
        string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nar-y-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Narrow the page width to make hyphenation more likely.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Add a paragraph with a long word.
        builder.Font.Size = 14;
        builder.Writeln("The word extraordinarycharacteristically demonstrates hyphenation.");

        // Insert a table with a single cell containing the same long word.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("The word extraordinarycharacteristically inside a table cell.");

        // Suppress hyphenation for the paragraph inside this cell.
        Paragraph cellParagraph = table.FirstRow.FirstCell.FirstParagraph;
        cellParagraph.ParagraphFormat.SuppressAutoHyphens = true;

        builder.EndTable();

        // Save the document.
        string outputPath = "HyphenationGlobalAndCell.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Clean up the temporary dictionary file.
        if (File.Exists(dictPath))
            File.Delete(dictPath);
    }
}
