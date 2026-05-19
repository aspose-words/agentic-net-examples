using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for dictionary and output files.
        const string dictionaryPath = "hyph_en_US.dic";
        const string outputBefore = "hyphenated_before.pdf";
        const string outputAfter = "hyphenated_after.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create an initial Hunspell dictionary with basic patterns.
        // -----------------------------------------------------------------
        // The dictionary format: first line is the encoding, subsequent lines are
        // word=hyphenation-pattern (hyphens indicate allowed break points).
        string initialDictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n" +
            // Note: No pattern for the technical term \"microprocessor\" yet.
            "microprocessor=microprocessor\n";

        File.WriteAllText(dictionaryPath, initialDictionaryContent);

        // Register the dictionary for the \"en-US\" locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // --------------------------------------------------------------
        // Step 2: Build a sample document that contains long words.
        // --------------------------------------------------------------
        Document docBefore = new Document();
        DocumentBuilder builder = new DocumentBuilder(docBefore);

        // Narrow the page width to force line wrapping and hyphenation.
        docBefore.FirstSection.PageSetup.PageWidth = 200;
        docBefore.FirstSection.PageSetup.LeftMargin = 20;
        docBefore.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph with technical terminology.
        builder.Font.Size = 12;
        builder.Writeln(
            "The extraordinarycharacteristically advanced communication " +
            "systems rely on the microprocessor to perform complex calculations.");

        // Enable automatic hyphenation.
        docBefore.HyphenationOptions.AutoHyphenation = true;

        // Save the document before adding the custom pattern.
        docBefore.Save(outputBefore);
        if (!File.Exists(outputBefore))
            throw new InvalidOperationException($"Failed to create '{outputBefore}'.");

        // --------------------------------------------------------------
        // Step 3: Update the dictionary with a custom hyphenation pattern.
        // --------------------------------------------------------------
        // Append a proper hyphenation pattern for the technical term.
        string updatedPattern = "microprocessor=micro-pro-cessor\n";
        File.AppendAllText(dictionaryPath, updatedPattern);

        // Unregister the old dictionary and register the updated one.
        Hyphenation.UnregisterDictionary("en-US");
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // --------------------------------------------------------------
        // Step 4: Build a new document using the updated dictionary.
        // --------------------------------------------------------------
        Document docAfter = new Document();
        DocumentBuilder builderAfter = new DocumentBuilder(docAfter);

        // Apply the same page setup.
        docAfter.FirstSection.PageSetup.PageWidth = 200;
        docAfter.FirstSection.PageSetup.LeftMargin = 20;
        docAfter.FirstSection.PageSetup.RightMargin = 20;

        // Write the same paragraph.
        builderAfter.Font.Size = 12;
        builderAfter.Writeln(
            "The extraordinarycharacteristically advanced communication " +
            "systems rely on the microprocessor to perform complex calculations.");

        // Enable automatic hyphenation.
        docAfter.HyphenationOptions.AutoHyphenation = true;

        // Save the document after the dictionary update.
        docAfter.Save(outputAfter);
        if (!File.Exists(outputAfter))
            throw new InvalidOperationException($"Failed to create '{outputAfter}'.");

        // Clean up: optional removal of temporary files.
        // File.Delete(dictionaryPath);
    }
}
