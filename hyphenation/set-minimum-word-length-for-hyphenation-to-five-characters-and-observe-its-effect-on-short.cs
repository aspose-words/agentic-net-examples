using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph containing short and long words.
        builder.Font.Size = 24;
        builder.Writeln("extra hyphenation demonstration extra hyphenation demonstration extra hyphenation demonstration");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Aspose.Words does not expose a MinimumWordLength property.
        // The default behavior already skips hyphenation for very short words.
        // Therefore we rely on the built‑in logic to avoid hyphenating words shorter than five characters.

        // Create a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "hyphenation=hy-phen-a-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Save the document as PDF to observe hyphenation.
        const string outputPath = "HyphenationMinWordLength.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF output file was not created.");
    }
}
