using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
@"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion");

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        // The Hyphenation class resides directly in the Aspose.Words namespace.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation and adjust the hyphenation zone to reduce large gaps.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 180; // 0.125 inch (default is 360)

        // Use compressed justification to tighten spacing after hyphenation.
        doc.JustificationMode = JustificationMode.Compress;

        // Set the paragraph to justified alignment.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;

        // Ensure the text uses the locale that matches the registered dictionary.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 12;

        // Add a long paragraph that will be hyphenated.
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication.");

        // Save the document to PDF so the layout can be inspected.
        const string outputFile = "AdjustedJustification.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException($"Expected output file '{outputFile}' was not created.");
    }
}
