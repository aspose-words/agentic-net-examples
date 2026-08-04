using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationExample
{
    public static void Main()
    {
        // Paths for the temporary files used in the example.
        const string sourceDocPath = "source.docx";
        const string hyphenationDictPath = "hyph_en_US.dic";
        const string outputDocPath = "hyphenated.docx";

        // 1. Create a minimal hyphenation dictionary for English (US).
        // The first line must be the encoding identifier, followed by word=pattern lines.
        // The patterns use hyphens to indicate allowed break points.
        File.WriteAllText(hyphenationDictPath,
            "UTF-8\n" +
            "hyphenation=hy-phen-ation\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can apply hyphenation.
        Hyphenation.RegisterDictionary("en-US", hyphenationDictPath);

        // 2. Build a sample document that contains long words which can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping and thus hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the language of the text to match the registered dictionary.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication hyphenation demonstration");

        // Save the source document (required because the task specifies loading an existing DOCX).
        doc.Save(sourceDocPath);

        // 3. Load the previously saved document.
        Document loadedDoc = new Document(sourceDocPath);

        // 4. Enable automatic hyphenation.
        loadedDoc.HyphenationOptions.AutoHyphenation = true;
        loadedDoc.HyphenationOptions.HyphenateCaps = true; // Hyphenate words in all caps if any.
        loadedDoc.HyphenationOptions.ConsecutiveHyphenLimit = 2; // Example setting.

        // Force layout update so that hyphenation is applied before saving.
        loadedDoc.UpdatePageLayout();

        // 5. Save the hyphenated document.
        loadedDoc.Save(outputDocPath);

        // 6. Validate that the output file was created.
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException($"The expected output file '{outputDocPath}' was not created.");

        // Clean up temporary files (optional).
        // Comment out the following lines if you wish to inspect the generated files.
        File.Delete(sourceDocPath);
        File.Delete(hyphenationDictPath);
        // The hyphenated document is left on disk as the final result.
    }
}
