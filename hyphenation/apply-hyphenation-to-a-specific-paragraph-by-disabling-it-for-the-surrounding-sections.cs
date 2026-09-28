using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (en-US).
        const string dictionaryFileName = "hyph_en_US.dic";
        const string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-na-ry-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n";

        File.WriteAllText(dictionaryFileName, dictionaryContent);

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictionaryFileName);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to force line wrapping and hyphenation.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200; // points
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Set the locale for hyphenation.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Configure hyphenation options globally.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Paragraph 1 – hyphenation disabled.
        doc.HyphenationOptions.AutoHyphenation = false;
        builder.Writeln("extraordinarycharacteristically communication internationalization");

        // Paragraph 2 – hyphenation enabled.
        doc.HyphenationOptions.AutoHyphenation = true;
        builder.Writeln("extraordinarycharacteristically communication internationalization");

        // Paragraph 3 – hyphenation disabled again.
        doc.HyphenationOptions.AutoHyphenation = false;
        builder.Writeln("extraordinarycharacteristically communication internationalization");

        // Save the document as PDF.
        const string outputFileName = "HyphenationDemo.pdf";
        doc.Save(outputFileName, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFileName))
        {
            throw new InvalidOperationException($"The expected output file '{outputFileName}' was not created.");
        }

        // Clean up the temporary dictionary file.
        if (File.Exists(dictionaryFileName))
        {
            File.Delete(dictionaryFileName);
        }
    }
}
