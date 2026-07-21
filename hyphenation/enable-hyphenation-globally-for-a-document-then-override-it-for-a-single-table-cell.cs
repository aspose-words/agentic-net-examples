using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the hyphenation dictionary and the output document.
        const string dictPath = "hyph_en_US.dic";
        const string outputPath = "HyphenationExample.docx";

        // Create a minimal hyphenation dictionary for English (US).
        // The dictionary format: first line is the encoding, followed by word=hyphenation-points.
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so Aspose.Words can hyphenate words in this language.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the locale of the text to English (US) to match the dictionary.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 12;

        // Write a paragraph with a long word that can be hyphenated.
        builder.Writeln("extraordinarycharacteristically");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Insert a table with a single cell containing the same long word.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("extraordinarycharacteristically");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the first cell of the table.
        Cell firstCell = (Cell)doc.GetChild(NodeType.Cell, 0, true);
        if (firstCell == null)
            throw new InvalidOperationException("Failed to locate the table cell.");

        // Suppress hyphenation for the paragraph inside this cell.
        Paragraph cellParagraph = firstCell.FirstParagraph;
        if (cellParagraph == null)
            throw new InvalidOperationException("The cell does not contain a paragraph.");
        cellParagraph.ParagraphFormat.SuppressAutoHyphens = true;

        // Save the document.
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");
    }
}
