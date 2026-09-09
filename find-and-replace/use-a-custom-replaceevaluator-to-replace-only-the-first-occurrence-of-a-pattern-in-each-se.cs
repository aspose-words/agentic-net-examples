using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two sections, each containing the word "PLACEHOLDER" three times.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // First section
        builder.Writeln("Section 1 - First occurrence: PLACEHOLDER");
        builder.Writeln("Section 1 - Second occurrence: PLACEHOLDER");
        builder.Writeln("Section 1 - Third occurrence: PLACEHOLDER");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section
        builder.Writeln("Section 2 - First occurrence: PLACEHOLDER");
        builder.Writeln("Section 2 - Second occurrence: PLACEHOLDER");
        builder.Writeln("Section 2 - Third occurrence: PLACEHOLDER");

        // Save the source document.
        const string inputPath = "Input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        var loadedDoc = new Document(inputPath);

        // Set up find‑replace options with a custom callback.
        var options = new FindReplaceOptions(new FirstOccurrencePerSectionCallback())
        {
            // Ensure the search is case‑sensitive (optional).
            MatchCase = true
        };

        // Replace only the first "PLACEHOLDER" in each section with "REPLACED".
        int replacedCount = loadedDoc.Range.Replace("PLACEHOLDER", "REPLACED", options);

        // Validate that replacements were made (at least one per section).
        if (replacedCount == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document.
        const string outputPath = "Output.docx";
        loadedDoc.Save(outputPath);

        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to '{outputPath}'.");
    }

    /// <summary>
    /// Replaces only the first match of the search pattern in each section.
    /// Subsequent matches within the same section are skipped.
    /// </summary>
    private class FirstOccurrencePerSectionCallback : IReplacingCallback
    {
        // Tracks whether a replacement has already occurred in a given section.
        private readonly HashSet<Section> _sectionsReplaced = new HashSet<Section>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the Section that contains the current match.
            Node? node = args.MatchNode;
            while (node != null && !(node is Section))
                node = node.ParentNode;

            if (node is not Section section)
                return ReplaceAction.Skip; // Safety fallback.

            // If this section has not been replaced yet, allow the replacement.
            if (_sectionsReplaced.Add(section))
            {
                // The replacement text is already supplied via the Range.Replace call,
                // but we can modify it here if needed.
                // args.Replacement = "REPLACED";
                return ReplaceAction.Replace;
            }

            // Skip all further matches in this section.
            return ReplaceAction.Skip;
        }
    }
}
