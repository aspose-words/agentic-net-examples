using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class HyphenationExample
{
    public static void Main()
    {
        // Prepare hyphenation dictionaries for English (US) and German (Switzerland).
        const string enDictPath = "hyph_en_US.dic";
        const string deDictPath = "hyph_de_CH.dic";

        // Minimal valid dictionary content.
        File.WriteAllText(enDictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        File.WriteAllText(deDictPath,
            "UTF-8\n" +
            "internationalisierung=in-ter-na-tion-alisie-rung\n" +
            "kommunikation=kom-mu-ni-ka-tion\n" +
            "extraordinaer=ex-tra-or-di-naer\n");

        // Register the dictionaries.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", enDictPath);
        Aspose.Words.Hyphenation.RegisterDictionary("de-CH", deDictPath);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // ---------- Section 1: English ----------
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Insert a section break to start the next language.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2: German ----------
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;
        builder.Writeln("extraordinaer internationalisierung kommunikation");

        // Save the document to PDF to visualize hyphenation.
        const string outputPath = "HyphenatedOutput.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
