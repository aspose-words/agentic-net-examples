using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for output and temporary dictionary files.
        const string outputPdf = "HyphenatedOutput.pdf";
        const string dictEnPath = "hyph_en_US.dic";
        const string dictDePath = "hyph_de_CH.dic";

        // Create minimal hyphenation dictionaries required for the example.
        // The first line must be the encoding identifier.
        File.WriteAllText(dictEnPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        File.WriteAllText(dictDePath,
            "UTF-8\n" +
            "ausgezeichnete=aus-ge-zeich-net-te\n" +
            "kommunikation=ko-mmu-ni-ka-tion\n" +
            "internationalisierung=in-ter-na-tio-na-li-sie-rung\n");

        // Register the dictionaries for the corresponding language codes.
        Hyphenation.RegisterDictionary("en-US", dictEnPath);
        Hyphenation.RegisterDictionary("de-CH", dictDePath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // ---------- Section 1 – English ----------
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Insert a section break to start a new section with a different language.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Section 2 – German ----------
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;
        builder.Writeln("ausgezeichnete kommunikation und internationalisierung");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 360; // default
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document to PDF.
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The expected PDF output was not created.");

        // Clean up temporary dictionary files (optional).
        File.Delete(dictEnPath);
        File.Delete(dictDePath);
    }
}
