using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string inputDocPath = "source.docx";
        const string dictionaryPath = "hyph_en_US.dic";
        const string outputDocPath = "hyphenated_output.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document with long words that can be hyphenated
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping and hyphenation
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the document language to English (US)
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Write a paragraph containing words that will be hyphenated
        builder.Font.Size = 12;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "hyperresponsibility misunderstanding incomprehensibilities " +
            "characteristically extraordinarycharacteristically");

        // -----------------------------------------------------------------
        // 2. Create a minimal hyphenation dictionary file locally
        // -----------------------------------------------------------------
        // The dictionary format: first line is "UTF-8", subsequent lines are
        // word=hyphenation-points (hyphens separate syllable fragments)
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nar-y-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n" +
            "hyperresponsibility=hy-per-re-spon-si-bi-li-ty\n" +
            "misunderstanding=mis-un-der-stand-ing\n" +
            "incomprehensibilities=in-com-pre-hen-si-bi-li-ties\n" +
            "characteristically=char-ac-ter-is-ti-cal-ly\n";

        File.WriteAllText(dictionaryPath, dictionaryContent);

        // -----------------------------------------------------------------
        // 3. Register the dictionary and enable automatic hyphenation
        // -----------------------------------------------------------------
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Enable automatic hyphenation for the document
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Force layout rebuild so hyphenation is applied before saving
        doc.UpdatePageLayout();

        // -----------------------------------------------------------------
        // 4. Save the document preserving hyphenation marks
        // -----------------------------------------------------------------
        doc.Save(outputDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created
        // -----------------------------------------------------------------
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException($"The expected output file '{outputDocPath}' was not created.");

        // Clean up temporary files (optional)
        // File.Delete(dictionaryPath);
        // File.Delete(inputDocPath);
    }
}
