using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names for the dictionary and the resulting PDF.
        const string dictionaryFileName = "hyph_en_US.dic";
        const string outputPdf = "hyphenated.pdf";

        // Create a minimal Hunspell hyphenation dictionary.
        // The first line is the encoding identifier (e.g., "UTF-8").
        // Subsequent lines contain word=hyphenated-pattern entries.
        // Include a custom technical term "microprocessor".
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n" +
            "microprocessor=mi-cro-pro-cess-or\n";

        // Write the dictionary to the local file system.
        File.WriteAllText(dictionaryFileName, dictionaryContent);

        // Register the dictionary for the "en-US" locale.
        // Use the static RegisterDictionary method of Aspose.Words.Hyphenation.
        Hyphenation.RegisterDictionary("en-US", dictionaryFileName);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the font locale before writing text so that the runs inherit the correct locale.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 12;

        // Write a paragraph containing long words that can be hyphenated.
        builder.Writeln(
            "The extraordinarycharacteristically internationalization communication " +
            "process often involves complex microprocessor architectures that " +
            "require careful analysis.");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to PDF.
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The expected PDF output was not created.");
    }
}
