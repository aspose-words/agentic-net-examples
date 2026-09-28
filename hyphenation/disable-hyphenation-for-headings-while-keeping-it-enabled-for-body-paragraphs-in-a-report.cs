using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary with Aspose.Words for the "en-US" language.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a new document and a builder.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // ----- Heading (hyphenation disabled) -----
        // Use a language for which no dictionary is registered (e.g., French).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Font.LocaleId = (int)CultureInfo.GetCultureInfo("fr-FR").LCID;
        builder.Writeln("Extraordinarycharacteristically Internationalization Communication");

        // ----- Body paragraph (hyphenation enabled) -----
        // Use the language that has the registered dictionary (English US).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Font.LocaleId = (int)CultureInfo.GetCultureInfo("en-US").LCID;
        builder.Writeln("Extraordinarycharacteristically Internationalization Communication");

        // Save the document as PDF.
        const string outputPath = "Report.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output file was not created.");
    }
}
