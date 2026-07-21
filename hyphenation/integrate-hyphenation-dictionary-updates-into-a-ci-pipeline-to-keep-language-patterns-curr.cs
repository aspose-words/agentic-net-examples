using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the dictionary and output files.
        const string dictionaryPath = "hyph_en_US.dic";
        const string outputV1 = "hyphenated_v1.pdf";
        const string outputV2 = "hyphenated_v2.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create an initial hyphenation dictionary.
        // -----------------------------------------------------------------
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the en-US locale.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // -----------------------------------------------------------------
        // Step 2: Build a sample document that will be hyphenated.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the language of the text.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 24;

        // Write text containing words defined in the dictionary.
        builder.Writeln(
            "extraordinarycharacteristically communication extraordinarycharacteristically communication");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the first version of the document.
        doc.Save(outputV1);
        ValidateFileExists(outputV1, "First PDF output");

        // -----------------------------------------------------------------
        // Step 3: Simulate a CI pipeline update – modify the dictionary.
        // -----------------------------------------------------------------
        // Append a new pattern to the dictionary (e.g., for "internationalization").
        File.AppendAllText(dictionaryPath,
            "internationalization=inter-na-tion-al-i-za-tion\n");

        // Unregister the old dictionary and register the updated one.
        Aspose.Words.Hyphenation.UnregisterDictionary("en-US");
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Rebuild the layout to apply the updated hyphenation rules.
        doc.UpdatePageLayout();

        // Save the document again with the updated dictionary.
        doc.Save(outputV2);
        ValidateFileExists(outputV2, "Second PDF output");

        // Indicate successful completion.
        Console.WriteLine("Hyphenation dictionary update simulation completed successfully.");
    }

    private static void ValidateFileExists(string path, string description)
    {
        if (!File.Exists(path))
            throw new InvalidOperationException($"{description} was not created at '{path}'.");
    }
}
