using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class HyphenationExample
{
    public static void Main()
    {
        // Paths for temporary dictionary files.
        const string enDictPath = "hyph_en_US.dic";
        const string deDictPath = "hyph_de_CH.dic";

        // Create minimal English hyphenation dictionary.
        File.WriteAllText(enDictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Create minimal German hyphenation dictionary.
        File.WriteAllText(deDictPath,
            "UTF-8\n" +
            "Beispielwortzusammensetzung=Beis-piel-wort-zu-sam-men-set-zung\n");

        // Register both dictionaries.
        Hyphenation.RegisterDictionary("en-US", enDictPath);
        Hyphenation.RegisterDictionary("de-CH", deDictPath);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // English paragraph.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // New section for German text.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;
        builder.Writeln("Beispielwortzusammensetzung");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the result.
        const string outputPath = "HyphenatedOutput.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected output file was not created.");

        // Clean up temporary dictionary files.
        File.Delete(enDictPath);
        File.Delete(deDictPath);
    }
}
