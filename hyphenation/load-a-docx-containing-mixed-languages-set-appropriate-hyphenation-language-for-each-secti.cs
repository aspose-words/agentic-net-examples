using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class HyphenationExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // ---------- English section ----------
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        // Insert a section break to start a new language section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- German section ----------
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;
        builder.Writeln("Donaudampfschifffahrtsgesellschaftskapitän");
        builder.Writeln("Entwicklungszusammenarbeit");

        // ---------- Create minimal hyphenation dictionaries ----------
        // English (en-US) dictionary.
        string enDictPath = "hyph_en_US.dic";
        File.WriteAllText(enDictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-na-ry-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // German (de-CH) dictionary.
        string deDictPath = "hyph_de_CH.dic";
        File.WriteAllText(deDictPath,
            "UTF-8\n" +
            "Donaudampfschifffahrtsgesellschaftskapitän=Do-nau-dampf-schiff-fahrts-ge-sell-schafts-ka-pi-tän\n" +
            "Entwicklungszusammenarbeit=Ent-wick-lungs-zu-sam-men-ar-beit\n");

        // Register the dictionaries.
        Hyphenation.RegisterDictionary("en-US", enDictPath);
        Hyphenation.RegisterDictionary("de-CH", deDictPath);

        // Verify registration.
        if (!Hyphenation.IsDictionaryRegistered("en-US") || !Hyphenation.IsDictionaryRegistered("de-CH"))
            throw new InvalidOperationException("Failed to register hyphenation dictionaries.");

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the result as PDF to visualize hyphenation.
        string outputPath = "Hyphenated.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output file was not created.");
    }
}
