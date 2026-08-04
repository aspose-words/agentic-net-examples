using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a minimal Arabic hyphenation dictionary.
        const string dictPath = "hyph_ar_SA.dic";
        const string dictContent =
            "UTF-8\n" +
            "البرمجة=ال-برم-جة\n" +
            "التطوير=الت-طو-ير\n" +
            "المستقبل=المست-قبل\n";

        File.WriteAllText(dictPath, dictContent);

        // Register the dictionary for the Arabic (Saudi Arabia) locale.
        Aspose.Words.Hyphenation.RegisterDictionary("ar-SA", dictPath);

        // Create a new document and configure its layout.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the font locale to Arabic.
        builder.Font.LocaleId = new CultureInfo("ar-SA").LCID;
        builder.Font.Name = "Arial";

        // Add Arabic text that will require hyphenation when wrapped.
        builder.Writeln("البرمجة والتطوير هما مفتاح المستقبل في عالم التكنولوجيا.");

        // Narrow the page width to force line wrapping and hyphenation.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200; // points
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Enable hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

        // Save the document as PDF.
        const string outputPath = "ArabicHyphenated.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF output file was not created.");
    }
}
