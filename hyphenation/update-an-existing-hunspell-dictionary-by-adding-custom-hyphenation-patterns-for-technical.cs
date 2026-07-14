using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the hyphenation dictionary.
        const string dictFile = "hyph_en_US.dic";

        // -----------------------------------------------------------------
        // 1. Create an initial Hunspell dictionary with a single pattern.
        // -----------------------------------------------------------------
        File.WriteAllText(dictFile,
            "UTF-8\n" +
            "technologies=tech-no-lo-gi-es\n"); // pattern for the word "technologies"

        // Register the dictionary for English (US).
        Hyphenation.RegisterDictionary("en-US", dictFile);

        // -----------------------------------------------------------------
        // 2. Build a document that contains technical terminology.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width forces line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the locale to match the registered dictionary.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Write sample text containing words that will be hyphenated.
        builder.Writeln(
            "The field of technologies evolves rapidly. " +
            "Understanding microprocessor architecture is essential.");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the first PDF (uses the initial dictionary).
        const string initialPdf = "hyphenated_initial.pdf";
        doc.Save(initialPdf);
        if (!File.Exists(initialPdf))
            throw new InvalidOperationException("Initial PDF was not created.");

        // -----------------------------------------------------------------
        // 3. Update the dictionary with an additional pattern.
        // -----------------------------------------------------------------
        // Append a new hyphenation rule for "microprocessor".
        File.AppendAllText(dictFile, "microprocessor=micro-pro-cess-or\n");

        // Re‑register the updated dictionary.
        Hyphenation.UnregisterDictionary("en-US");
        Hyphenation.RegisterDictionary("en-US", dictFile);

        // After changing the dictionary we need to rebuild the layout.
        doc.UpdatePageLayout();

        // Save a second PDF that reflects the updated dictionary.
        const string updatedPdf = "hyphenated_updated.pdf";
        doc.Save(updatedPdf);
        if (!File.Exists(updatedPdf))
            throw new InvalidOperationException("Updated PDF was not created.");
    }
}
