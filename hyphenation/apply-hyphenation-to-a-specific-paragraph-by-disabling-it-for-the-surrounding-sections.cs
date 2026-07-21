using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationExample
{
    public static void Main()
    {
        // Path for the hyphenation dictionary.
        const string dictionaryPath = "hyph_en_US.dic";

        // Create a minimal dictionary file required for the example.
        // The dictionary format: first line is the encoding, subsequent lines are word=hyphenation-points.
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nar-y-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the English (United States) locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Narrow the page width so that long words will need to wrap (and thus hyphenate).
        doc.FirstSection.PageSetup.PageWidth = 300;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Use the same locale for all text to match the registered dictionary.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // First paragraph – hyphenation disabled.
        builder.ParagraphFormat.SuppressAutoHyphens = true;
        builder.Writeln("This paragraph has hyphenation disabled. extraordinarycharacteristically communication.");

        // Second paragraph – hyphenation enabled (default value is false, so we set it explicitly).
        builder.ParagraphFormat.SuppressAutoHyphens = false;
        builder.Writeln("This paragraph has hyphenation enabled. extraordinarycharacteristically communication.");

        // Third paragraph – hyphenation disabled again.
        builder.ParagraphFormat.SuppressAutoHyphens = true;
        builder.Writeln("Again hyphenation disabled. extraordinarycharacteristically communication.");

        // Save the document to PDF so that hyphenation can be observed in the rendered output.
        const string outputPath = "HyphenationExample.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The expected output file '{outputPath}' was not created.");
    }
}
