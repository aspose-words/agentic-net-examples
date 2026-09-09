using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so Aspose.Words can hyphenate the words above.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph with long words that can be hyphenated.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Insert a table with a single cell containing the same text.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the first cell of the first table.
        Table table = doc.FirstSection.Body.Tables[0];
        Cell cell = table.FirstRow.FirstCell;

        // Suppress hyphenation for all paragraphs inside this cell.
        foreach (Paragraph para in cell.Paragraphs)
        {
            para.ParagraphFormat.SuppressAutoHyphens = true;
        }

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to PDF (any format works; PDF shows hyphenation clearly).
        const string outputFileName = "HyphenationExample.pdf";
        doc.Save(outputFileName, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputFileName))
            throw new InvalidOperationException("The output PDF was not created.");
    }
}
