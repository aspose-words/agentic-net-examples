using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1
        builder.Writeln("Section 1 - First occurrence: TARGET");
        builder.Writeln("Section 1 - Second occurrence: TARGET");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - First occurrence: TARGET");
        builder.Writeln("Section 2 - Second occurrence: TARGET");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - First occurrence: TARGET");
        builder.Writeln("Section 3 - Second occurrence: TARGET");

        // Save the source document (optional, just to demonstrate lifecycle).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Prepare the replace callback that limits replacement to the first match per section.
        var callback = new FirstOccurrencePerSectionReplacer();

        // Configure find/replace options.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback,
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replacement using a regular expression that matches the word "TARGET".
        int replacedCount = loaded.Range.Replace(new Regex(@"\bTARGET\b"), "REPLACED", options);

        // Validate that at least one replacement was made in each section.
        if (callback.ReplacementsPerSection.Count != loaded.Sections.Count)
            throw new InvalidOperationException("Expected a replacement in each section.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);

        // Simple verification output (no interactive prompts).
        Console.WriteLine($"Total replacements performed: {replacedCount}");
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }

    // Callback that replaces only the first occurrence of a match within each section.
    private class FirstOccurrencePerSectionReplacer : IReplacingCallback
    {
        // Tracks whether a replacement has already occurred for a given section.
        public Dictionary<Section, bool> ReplacementsPerSection { get; } = new Dictionary<Section, bool>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the section that contains the current match.
            Section section = args.MatchNode.GetAncestor(NodeType.Section) as Section;
            if (section == null)
                return ReplaceAction.Skip; // Safety check.

            // If we have not replaced anything in this section yet, allow replacement.
            if (!ReplacementsPerSection.ContainsKey(section))
            {
                ReplacementsPerSection[section] = true;
                return ReplaceAction.Replace;
            }

            // Otherwise skip this match.
            return ReplaceAction.Skip;
        }
    }
}
