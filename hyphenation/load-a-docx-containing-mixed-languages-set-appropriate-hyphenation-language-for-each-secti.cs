using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationExample
{
    public static void Main()
    {
        // Create minimal hyphenation dictionaries for English (en-US) and German (de-DE).
        const string enDictPath = "hyph_en_US.dic";
        const string deDictPath = "hyph_de_DE.dic";

        File.WriteAllText(enDictPath,
@"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
communication=com-mu-ni-ca-tion
internationalization=in-ter-na-tion-al-i-za-tion");

        File.WriteAllText(deDictPath,
@"UTF-8
unabhaengigkeitserklaerung=un-ab-haa-ngi-keit-ser-klä-rung
kommunikation=ko-mmu-ni-ka-tion
internationalisierung=in-ter-na-tio-na-li-sie-rung");

        // Register the dictionaries with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", enDictPath);
        Hyphenation.RegisterDictionary("de-DE", deDictPath);

        // Build a sample DOCX containing mixed‑language sections.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // English section.
        builder.Font.LocaleId = CultureInfo.GetCultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically communication internationalization");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // German section.
        builder.Font.LocaleId = CultureInfo.GetCultureInfo("de-DE").LCID;
        builder.Writeln("unabhaengigkeitserklaerung kommunikation internationalisierung");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 5; // points

        // Save the document as PDF to visualize hyphenation.
        const string outputPdf = "mixed_hyphenated.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException($"Expected output file '{outputPdf}' was not created.");
        }

        // Optional cleanup of temporary dictionary files.
        // File.Delete(enDictPath);
        // File.Delete(deDictPath);
    }
}
