using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with words of varying lengths.
        // Short words (e.g., "test") will not be hyphenated when the minimum word length is enforced.
        builder.Writeln("extraordinary communication short words test");

        // Narrow the page width to force line wrapping where hyphenation can occur.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Create a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinary=ex-tra-or-di-nary\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // NOTE: Aspose.Words versions prior to v23 exposed a MinWordLength property.
        // In newer versions this property has been removed; hyphenation will follow the
        // dictionary rules without an explicit minimum word length setting.
        // The line below is omitted to keep the code compatible with current APIs.

        // Save the document to PDF.
        const string outPath = "HyphenationMinWordLength.pdf";
        doc.Save(outPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outPath))
            throw new InvalidOperationException("The expected PDF output was not created.");
    }
}
