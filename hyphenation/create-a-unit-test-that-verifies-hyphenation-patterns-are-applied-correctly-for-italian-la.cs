using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using static Aspose.Words.Hyphenation; // Allows direct access to RegisterDictionary, IsDictionaryRegistered, etc.

public class HyphenationItalianTest
{
    public static void Main()
    {
        // Define temporary file names.
        const string dictPath = "hyph_it_IT.dic";
        const string outputPath = "Hyphenated_it.pdf";

        // Clean up any previous runs.
        if (File.Exists(dictPath)) File.Delete(dictPath);
        if (File.Exists(outputPath)) File.Delete(outputPath);

        // Create a minimal Italian hyphenation dictionary.
        // First line: encoding. Subsequent lines: word=pattern (hyphens mark break points).
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinariamente=extra-or-di-nar-mente\n" +
            "incredibilmente=in-cred-i-bil-men-te\n";

        File.WriteAllText(dictPath, dictionaryContent);

        // Register the dictionary for the Italian locale.
        RegisterDictionary("it-IT", dictPath);

        // Verify registration succeeded.
        if (!IsDictionaryRegistered("it-IT"))
            throw new InvalidOperationException("Italian hyphenation dictionary was not registered.");

        // Create a new document and enable automatic hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        doc.HyphenationOptions.AutoHyphenation = true;

        // Narrow page width forces line wrapping, making hyphenation visible.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 cm)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the font locale to Italian.
        builder.Font.LocaleId = new CultureInfo("it-IT").LCID;
        builder.Font.Size = 24;

        // Paragraph containing words that can be hyphenated.
        string italianText = "extraordinariamente incredibilmente " +
                             "extraordinariamente incredibilmente " +
                             "extraordinariamente incredibilmente.";
        builder.Writeln(italianText);

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected PDF output was not created.");

        // Ensure the hyphenation option remains enabled after saving.
        if (!doc.HyphenationOptions.AutoHyphenation)
            throw new InvalidOperationException("Auto hyphenation option was not preserved.");

        // Optional clean‑up (commented out to allow inspection of the files after run).
        // File.Delete(dictPath);
        // File.Delete(outputPath);
    }
}
